---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 07/26/2019
localization_priority: Priority
ms.openlocfilehash: 7039ca59af22f1101bdff7b6bcd4506497d6c9cd
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940838"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="ec7eb-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ec7eb-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="ec7eb-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="ec7eb-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="ec7eb-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="ec7eb-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="ec7eb-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="ec7eb-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="ec7eb-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="ec7eb-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="ec7eb-108">Excel</span><span class="sxs-lookup"><span data-stu-id="ec7eb-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ec7eb-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="ec7eb-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ec7eb-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="ec7eb-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ec7eb-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ec7eb-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ec7eb-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="ec7eb-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="ec7eb-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-114">- TaskPane</span></span><br><span data-ttu-id="ec7eb-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-115">
        - Content</span></span><br><span data-ttu-id="ec7eb-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="ec7eb-116">
        - Custom Functions</span></span><br><span data-ttu-id="ec7eb-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="ec7eb-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ec7eb-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec7eb-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec7eb-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec7eb-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec7eb-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec7eb-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ec7eb-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ec7eb-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-130">
        - BindingEvents</span></span><br><span data-ttu-id="ec7eb-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-131">
        - CompressedFile</span></span><br><span data-ttu-id="ec7eb-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-132">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-133">
        - File</span></span><br><span data-ttu-id="ec7eb-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-134">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-136">
        - Selection</span></span><br><span data-ttu-id="ec7eb-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-137">
        - Settings</span></span><br><span data-ttu-id="ec7eb-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-138">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-139">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-140">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-142">Office on Windows</span></span><br><span data-ttu-id="ec7eb-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-144">- TaskPane</span></span><br><span data-ttu-id="ec7eb-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-145">
        - Content</span></span><br><span data-ttu-id="ec7eb-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="ec7eb-146">
        - Custom Functions</span></span><br><span data-ttu-id="ec7eb-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="ec7eb-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ec7eb-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec7eb-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec7eb-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec7eb-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec7eb-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec7eb-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ec7eb-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ec7eb-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-160">
        - BindingEvents</span></span><br><span data-ttu-id="ec7eb-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-161">
        - CompressedFile</span></span><br><span data-ttu-id="ec7eb-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-162">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-163">
        - File</span></span><br><span data-ttu-id="ec7eb-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-164">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-166">
        - Selection</span></span><br><span data-ttu-id="ec7eb-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-167">
        - Settings</span></span><br><span data-ttu-id="ec7eb-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-168">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-169">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-170">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-172">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-172">Office 2019 on Windows</span></span><br><span data-ttu-id="ec7eb-173">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec7eb-174">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-174">- TaskPane</span></span><br><span data-ttu-id="ec7eb-175">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-175">
        - Content</span></span><br><span data-ttu-id="ec7eb-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ec7eb-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec7eb-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec7eb-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec7eb-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec7eb-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec7eb-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec7eb-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-187">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-188">
        - CompressedFile</span></span><br><span data-ttu-id="ec7eb-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-189">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-190">
        - File</span></span><br><span data-ttu-id="ec7eb-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-191">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-193">
        - Selection</span></span><br><span data-ttu-id="ec7eb-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-194">
        - Settings</span></span><br><span data-ttu-id="ec7eb-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-195">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-196">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-197">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-199">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-199">Office 2016 on Windows</span></span><br><span data-ttu-id="ec7eb-200">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec7eb-201">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-201">- TaskPane</span></span><br><span data-ttu-id="ec7eb-202">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-202">
        - Content</span></span></td>
    <td><span data-ttu-id="ec7eb-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ec7eb-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec7eb-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-206">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-207">
        - CompressedFile</span></span><br><span data-ttu-id="ec7eb-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-208">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-209">
        - File</span></span><br><span data-ttu-id="ec7eb-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-210">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-212">
        - Selection</span></span><br><span data-ttu-id="ec7eb-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-213">
        - Settings</span></span><br><span data-ttu-id="ec7eb-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-214">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-215">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-216">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-218">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-218">Office 2013 on Windows</span></span><br><span data-ttu-id="ec7eb-219">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec7eb-220">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-220">
        - TaskPane</span></span><br><span data-ttu-id="ec7eb-221">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="ec7eb-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec7eb-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec7eb-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-224">
        - BindingEvents</span></span><br><span data-ttu-id="ec7eb-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-225">
        - CompressedFile</span></span><br><span data-ttu-id="ec7eb-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-226">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-227">
        - File</span></span><br><span data-ttu-id="ec7eb-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-228">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-230">
        - Selection</span></span><br><span data-ttu-id="ec7eb-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-231">
        - Settings</span></span><br><span data-ttu-id="ec7eb-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-232">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-233">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-234">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-236">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="ec7eb-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="ec7eb-237">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ec7eb-238">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-238">- TaskPane</span></span><br><span data-ttu-id="ec7eb-239">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-239">
        - Content</span></span><br><span data-ttu-id="ec7eb-240">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="ec7eb-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ec7eb-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec7eb-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec7eb-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec7eb-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec7eb-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec7eb-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ec7eb-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec7eb-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-252">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-253">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-254">
        - File</span></span><br><span data-ttu-id="ec7eb-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-255">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-257">
        - Selection</span></span><br><span data-ttu-id="ec7eb-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-258">
        - Settings</span></span><br><span data-ttu-id="ec7eb-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-259">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-260">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-261">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-263">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-263">Office apps on Mac</span></span><br><span data-ttu-id="ec7eb-264">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ec7eb-265">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-265">- TaskPane</span></span><br><span data-ttu-id="ec7eb-266">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-266">
        - Content</span></span><br><span data-ttu-id="ec7eb-267">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="ec7eb-267">
        - Custom Functions</span></span><br><span data-ttu-id="ec7eb-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ec7eb-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec7eb-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec7eb-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec7eb-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec7eb-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec7eb-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ec7eb-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ec7eb-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-281">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-282">
        - CompressedFile</span></span><br><span data-ttu-id="ec7eb-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-283">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-284">
        - File</span></span><br><span data-ttu-id="ec7eb-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-285">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-287">
        - PdfFile</span></span><br><span data-ttu-id="ec7eb-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-288">
        - Selection</span></span><br><span data-ttu-id="ec7eb-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-289">
        - Settings</span></span><br><span data-ttu-id="ec7eb-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-290">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-291">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-292">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-294">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-294">Office 2019 for Mac</span></span><br><span data-ttu-id="ec7eb-295">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec7eb-296">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-296">- TaskPane</span></span><br><span data-ttu-id="ec7eb-297">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-297">
        - Content</span></span><br><span data-ttu-id="ec7eb-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ec7eb-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec7eb-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec7eb-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec7eb-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec7eb-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec7eb-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec7eb-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-309">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-310">
        - CompressedFile</span></span><br><span data-ttu-id="ec7eb-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-311">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-312">
        - File</span></span><br><span data-ttu-id="ec7eb-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-313">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-315">
        - PdfFile</span></span><br><span data-ttu-id="ec7eb-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-316">
        - Selection</span></span><br><span data-ttu-id="ec7eb-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-317">
        - Settings</span></span><br><span data-ttu-id="ec7eb-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-318">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-319">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-320">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-322">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-322">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="ec7eb-323">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec7eb-324">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-324">- TaskPane</span></span><br><span data-ttu-id="ec7eb-325">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-325">
        - Content</span></span></td>
    <td><span data-ttu-id="ec7eb-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ec7eb-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec7eb-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-329">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-330">
        - CompressedFile</span></span><br><span data-ttu-id="ec7eb-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-331">
        - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-332">
        - File</span></span><br><span data-ttu-id="ec7eb-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-333">
        - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-335">
        - PdfFile</span></span><br><span data-ttu-id="ec7eb-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-336">
        - Selection</span></span><br><span data-ttu-id="ec7eb-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-337">
        - Settings</span></span><br><span data-ttu-id="ec7eb-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-338">
        - TableBindings</span></span><br><span data-ttu-id="ec7eb-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-339">
        - TableCoercion</span></span><br><span data-ttu-id="ec7eb-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-340">
        - TextBindings</span></span><br><span data-ttu-id="ec7eb-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ec7eb-342">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="ec7eb-343">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="ec7eb-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ec7eb-344">Платформа</span><span class="sxs-lookup"><span data-stu-id="ec7eb-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ec7eb-345">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="ec7eb-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ec7eb-346">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ec7eb-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ec7eb-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-348">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="ec7eb-348">Office on the web</span></span></td>
    <td><span data-ttu-id="ec7eb-349">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="ec7eb-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ec7eb-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-351">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-351">Office on Windows</span></span><br><span data-ttu-id="ec7eb-352">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ec7eb-353">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="ec7eb-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ec7eb-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-355">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-355">Office for Mac</span></span><br><span data-ttu-id="ec7eb-356">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="ec7eb-357">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="ec7eb-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ec7eb-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="ec7eb-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="ec7eb-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec7eb-360">Платформа</span><span class="sxs-lookup"><span data-stu-id="ec7eb-360">Platform</span></span></th>
    <th><span data-ttu-id="ec7eb-361">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="ec7eb-361">Extension points</span></span></th>
    <th><span data-ttu-id="ec7eb-362">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ec7eb-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec7eb-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-364">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="ec7eb-364">Office on the web</span></span><br><span data-ttu-id="ec7eb-365">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-365">Modern</span></span></td>
    <td> <span data-ttu-id="ec7eb-366">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-366">- Mail Read</span></span><br><span data-ttu-id="ec7eb-367">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-367">
      - Mail Compose</span></span><br><span data-ttu-id="ec7eb-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec7eb-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ec7eb-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ec7eb-376">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-377">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="ec7eb-377">Office on the web</span></span><br><span data-ttu-id="ec7eb-378">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-378">(classic)</span></span></td>
    <td> <span data-ttu-id="ec7eb-379">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-379">- Mail Read</span></span><br><span data-ttu-id="ec7eb-380">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-380">
      - Mail Compose</span></span><br><span data-ttu-id="ec7eb-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec7eb-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ec7eb-388">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-389">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-389">Office on Windows</span></span><br><span data-ttu-id="ec7eb-390">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-391">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-391">- Mail Read</span></span><br><span data-ttu-id="ec7eb-392">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-392">
      - Mail Compose</span></span><br><span data-ttu-id="ec7eb-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ec7eb-394">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="ec7eb-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ec7eb-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec7eb-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ec7eb-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ec7eb-402">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-403">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-403">Office 2019 on Windows</span></span><br><span data-ttu-id="ec7eb-404">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-405">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-405">- Mail Read</span></span><br><span data-ttu-id="ec7eb-406">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-406">
      - Mail Compose</span></span><br><span data-ttu-id="ec7eb-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ec7eb-408">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="ec7eb-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ec7eb-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec7eb-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ec7eb-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ec7eb-416">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-417">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-417">Office 2016 on Windows</span></span><br><span data-ttu-id="ec7eb-418">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-419">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-419">- Mail Read</span></span><br><span data-ttu-id="ec7eb-420">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-420">
      - Mail Compose</span></span><br><span data-ttu-id="ec7eb-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ec7eb-422">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="ec7eb-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ec7eb-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ec7eb-427">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-428">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-428">Office 2013 on Windows</span></span><br><span data-ttu-id="ec7eb-429">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-430">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-430">- Mail Read</span></span><br><span data-ttu-id="ec7eb-431">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="ec7eb-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="ec7eb-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ec7eb-436">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-437">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="ec7eb-437">Office apps on iOS</span></span><br><span data-ttu-id="ec7eb-438">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-439">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-439">- Mail Read</span></span><br><span data-ttu-id="ec7eb-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ec7eb-446">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-447">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-447">Office apps on Mac</span></span><br><span data-ttu-id="ec7eb-448">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-449">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-449">- Mail Read</span></span><br><span data-ttu-id="ec7eb-450">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-450">
      - Mail Compose</span></span><br><span data-ttu-id="ec7eb-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec7eb-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ec7eb-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ec7eb-459">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-460">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-460">Office 2019 for Mac</span></span><br><span data-ttu-id="ec7eb-461">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-462">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-462">- Mail Read</span></span><br><span data-ttu-id="ec7eb-463">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-463">
      - Mail Compose</span></span><br><span data-ttu-id="ec7eb-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec7eb-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ec7eb-471">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-472">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-472">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="ec7eb-473">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-474">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-474">- Mail Read</span></span><br><span data-ttu-id="ec7eb-475">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-475">
      - Mail Compose</span></span><br><span data-ttu-id="ec7eb-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec7eb-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ec7eb-483">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-484">Office для Android</span><span class="sxs-lookup"><span data-stu-id="ec7eb-484">Office apps on Android</span></span><br><span data-ttu-id="ec7eb-485">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-486">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="ec7eb-486">- Mail Read</span></span><br><span data-ttu-id="ec7eb-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec7eb-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec7eb-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec7eb-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec7eb-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ec7eb-493">Недоступно</span><span class="sxs-lookup"><span data-stu-id="ec7eb-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="ec7eb-494">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="ec7eb-495">Word</span><span class="sxs-lookup"><span data-stu-id="ec7eb-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec7eb-496">Платформа</span><span class="sxs-lookup"><span data-stu-id="ec7eb-496">Platform</span></span></th>
    <th><span data-ttu-id="ec7eb-497">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="ec7eb-497">Extension points</span></span></th>
    <th><span data-ttu-id="ec7eb-498">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ec7eb-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec7eb-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-500">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="ec7eb-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="ec7eb-501">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-501">- TaskPane</span></span><br><span data-ttu-id="ec7eb-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-509">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-511">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-512">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-512">
         - File</span></span><br><span data-ttu-id="ec7eb-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-514">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-517">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-518">
         - Selection</span></span><br><span data-ttu-id="ec7eb-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-519">
         - Settings</span></span><br><span data-ttu-id="ec7eb-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-520">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-521">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-522">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-523">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-525">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-525">Office on Windows</span></span><br><span data-ttu-id="ec7eb-526">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-527">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-527">- TaskPane</span></span><br><span data-ttu-id="ec7eb-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-535">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-536">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-538">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-539">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-539">
         - File</span></span><br><span data-ttu-id="ec7eb-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-541">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-544">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-545">
         - Selection</span></span><br><span data-ttu-id="ec7eb-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-546">
         - Settings</span></span><br><span data-ttu-id="ec7eb-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-547">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-548">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-549">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-550">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-552">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-552">Office 2019 on Windows</span></span><br><span data-ttu-id="ec7eb-553">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-554">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-554">- TaskPane</span></span><br><span data-ttu-id="ec7eb-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-561">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-562">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-564">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-565">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-565">
         - File</span></span><br><span data-ttu-id="ec7eb-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-567">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-570">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-571">
         - Selection</span></span><br><span data-ttu-id="ec7eb-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-572">
         - Settings</span></span><br><span data-ttu-id="ec7eb-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-573">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-574">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-575">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-576">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-578">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-578">Office 2016 on Windows</span></span><br><span data-ttu-id="ec7eb-579">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-580">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ec7eb-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-584">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-585">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-587">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-588">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-588">
         - File</span></span><br><span data-ttu-id="ec7eb-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-590">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-593">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-594">
         - Selection</span></span><br><span data-ttu-id="ec7eb-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-595">
         - Settings</span></span><br><span data-ttu-id="ec7eb-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-596">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-597">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-598">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-599">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-601">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-601">Office 2013 on Windows</span></span><br><span data-ttu-id="ec7eb-602">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-603">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec7eb-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-606">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-607">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-609">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-610">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-610">
         - File</span></span><br><span data-ttu-id="ec7eb-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-612">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-615">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-616">
         - Selection</span></span><br><span data-ttu-id="ec7eb-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-617">
         - Settings</span></span><br><span data-ttu-id="ec7eb-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-618">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-619">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-620">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-621">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-623">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="ec7eb-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="ec7eb-624">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-625">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ec7eb-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-631">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-632">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-634">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-635">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-635">
         - File</span></span><br><span data-ttu-id="ec7eb-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-637">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-640">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-641">
         - Selection</span></span><br><span data-ttu-id="ec7eb-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-642">
         - Settings</span></span><br><span data-ttu-id="ec7eb-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-643">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-644">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-645">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-646">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-648">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-648">Office apps on Mac</span></span><br><span data-ttu-id="ec7eb-649">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-650">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-650">- TaskPane</span></span><br><span data-ttu-id="ec7eb-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="ec7eb-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-658">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-659">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-661">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-662">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-662">
         - File</span></span><br><span data-ttu-id="ec7eb-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-664">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-667">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-668">
         - Selection</span></span><br><span data-ttu-id="ec7eb-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-669">
         - Settings</span></span><br><span data-ttu-id="ec7eb-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-670">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-671">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-672">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-673">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-675">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-675">Office 2019 for Mac</span></span><br><span data-ttu-id="ec7eb-676">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-677">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-677">- TaskPane</span></span><br><span data-ttu-id="ec7eb-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec7eb-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec7eb-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ec7eb-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-684">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-685">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-687">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-688">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-688">
         - File</span></span><br><span data-ttu-id="ec7eb-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-690">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-693">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-694">
         - Selection</span></span><br><span data-ttu-id="ec7eb-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-695">
         - Settings</span></span><br><span data-ttu-id="ec7eb-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-696">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-697">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-698">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-699">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-701">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-701">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="ec7eb-702">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-703">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ec7eb-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-707">- BindingEvents</span></span><br><span data-ttu-id="ec7eb-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-708">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec7eb-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec7eb-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-710">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-711">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="ec7eb-711">
         - File</span></span><br><span data-ttu-id="ec7eb-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-713">
         - MatrixBindings</span></span><br><span data-ttu-id="ec7eb-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec7eb-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec7eb-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-716">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-717">
         - Selection</span></span><br><span data-ttu-id="ec7eb-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-718">
         - Settings</span></span><br><span data-ttu-id="ec7eb-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-719">
         - TableBindings</span></span><br><span data-ttu-id="ec7eb-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-720">
         - TableCoercion</span></span><br><span data-ttu-id="ec7eb-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec7eb-721">
         - TextBindings</span></span><br><span data-ttu-id="ec7eb-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-722">
         - TextCoercion</span></span><br><span data-ttu-id="ec7eb-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="ec7eb-724">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ec7eb-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ec7eb-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec7eb-726">Платформа</span><span class="sxs-lookup"><span data-stu-id="ec7eb-726">Platform</span></span></th>
    <th><span data-ttu-id="ec7eb-727">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="ec7eb-727">Extension points</span></span></th>
    <th><span data-ttu-id="ec7eb-728">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ec7eb-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec7eb-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-730">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="ec7eb-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="ec7eb-731">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-731">- Content</span></span><br><span data-ttu-id="ec7eb-732">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-732">
         - TaskPane</span></span><br><span data-ttu-id="ec7eb-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-738">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-738">- ActiveView</span></span><br><span data-ttu-id="ec7eb-739">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-739">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-740">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-741">
         - File</span></span><br><span data-ttu-id="ec7eb-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-742">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-743">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-743">
         - Selection</span></span><br><span data-ttu-id="ec7eb-744">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-744">
         - Settings</span></span><br><span data-ttu-id="ec7eb-745">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-745">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-746">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-746">Office on Windows</span></span><br><span data-ttu-id="ec7eb-747">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-747">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-748">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-748">- Content</span></span><br><span data-ttu-id="ec7eb-749">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-749">
         - TaskPane</span></span><br><span data-ttu-id="ec7eb-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-755">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-755">- ActiveView</span></span><br><span data-ttu-id="ec7eb-756">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-756">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-757">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-757">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-758">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-758">
         - File</span></span><br><span data-ttu-id="ec7eb-759">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-759">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-760">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-760">
         - Selection</span></span><br><span data-ttu-id="ec7eb-761">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-761">
         - Settings</span></span><br><span data-ttu-id="ec7eb-762">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-762">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-763">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-763">Office 2019 on Windows</span></span><br><span data-ttu-id="ec7eb-764">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-764">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-765">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-765">- Content</span></span><br><span data-ttu-id="ec7eb-766">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-766">
         - TaskPane</span></span><br><span data-ttu-id="ec7eb-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-770">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-770">- ActiveView</span></span><br><span data-ttu-id="ec7eb-771">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-771">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-772">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-772">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-773">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-773">
         - File</span></span><br><span data-ttu-id="ec7eb-774">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-774">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-775">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-775">
         - Selection</span></span><br><span data-ttu-id="ec7eb-776">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-776">
         - Settings</span></span><br><span data-ttu-id="ec7eb-777">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-777">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-778">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-778">Office 2016 on Windows</span></span><br><span data-ttu-id="ec7eb-779">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-779">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-780">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-780">- Content</span></span><br><span data-ttu-id="ec7eb-781">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-781">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec7eb-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-784">- ActiveView</span></span><br><span data-ttu-id="ec7eb-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-785">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-786">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-787">
         - File</span></span><br><span data-ttu-id="ec7eb-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-788">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-789">
         - Selection</span></span><br><span data-ttu-id="ec7eb-790">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-790">
         - Settings</span></span><br><span data-ttu-id="ec7eb-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-792">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-792">Office 2013 on Windows</span></span><br><span data-ttu-id="ec7eb-793">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-794">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-794">- Content</span></span><br><span data-ttu-id="ec7eb-795">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-795">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="ec7eb-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec7eb-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-798">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-798">- ActiveView</span></span><br><span data-ttu-id="ec7eb-799">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-799">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-800">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-800">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-801">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-801">
         - File</span></span><br><span data-ttu-id="ec7eb-802">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-802">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-803">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-803">
         - Selection</span></span><br><span data-ttu-id="ec7eb-804">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-804">
         - Settings</span></span><br><span data-ttu-id="ec7eb-805">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-805">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-806">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="ec7eb-806">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="ec7eb-807">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-807">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-808">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-808">- Content</span></span><br><span data-ttu-id="ec7eb-809">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-809">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-813">- ActiveView</span></span><br><span data-ttu-id="ec7eb-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-814">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-815">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-816">
         - File</span></span><br><span data-ttu-id="ec7eb-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-817">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-818">
         - Selection</span></span><br><span data-ttu-id="ec7eb-819">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-819">
         - Settings</span></span><br><span data-ttu-id="ec7eb-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-821">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-821">Office apps on Mac</span></span><br><span data-ttu-id="ec7eb-822">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-822">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec7eb-823">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-823">- Content</span></span><br><span data-ttu-id="ec7eb-824">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-824">
         - TaskPane</span></span><br><span data-ttu-id="ec7eb-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec7eb-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-830">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-830">- ActiveView</span></span><br><span data-ttu-id="ec7eb-831">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-831">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-832">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-832">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-833">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-833">
         - File</span></span><br><span data-ttu-id="ec7eb-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-834">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-835">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-835">
         - Selection</span></span><br><span data-ttu-id="ec7eb-836">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-836">
         - Settings</span></span><br><span data-ttu-id="ec7eb-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-838">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-838">Office 2019 for Mac</span></span><br><span data-ttu-id="ec7eb-839">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-840">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-840">- Content</span></span><br><span data-ttu-id="ec7eb-841">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-841">
         - TaskPane</span></span><br><span data-ttu-id="ec7eb-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-845">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-845">- ActiveView</span></span><br><span data-ttu-id="ec7eb-846">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-846">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-847">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-847">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-848">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-848">
         - File</span></span><br><span data-ttu-id="ec7eb-849">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-849">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-850">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-850">
         - Selection</span></span><br><span data-ttu-id="ec7eb-851">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-851">
         - Settings</span></span><br><span data-ttu-id="ec7eb-852">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-852">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-853">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-853">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="ec7eb-854">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-854">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-855">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-855">- Content</span></span><br><span data-ttu-id="ec7eb-856">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-856">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec7eb-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec7eb-859">- ActiveView</span></span><br><span data-ttu-id="ec7eb-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-860">
         - CompressedFile</span></span><br><span data-ttu-id="ec7eb-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-861">
         - DocumentEvents</span></span><br><span data-ttu-id="ec7eb-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec7eb-862">
         - File</span></span><br><span data-ttu-id="ec7eb-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec7eb-863">
         - PdfFile</span></span><br><span data-ttu-id="ec7eb-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-864">
         - Selection</span></span><br><span data-ttu-id="ec7eb-865">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-865">
         - Settings</span></span><br><span data-ttu-id="ec7eb-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-866">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ec7eb-867">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="ec7eb-867">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="ec7eb-868">OneNote</span><span class="sxs-lookup"><span data-stu-id="ec7eb-868">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec7eb-869">Платформа</span><span class="sxs-lookup"><span data-stu-id="ec7eb-869">Platform</span></span></th>
    <th><span data-ttu-id="ec7eb-870">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="ec7eb-870">Extension points</span></span></th>
    <th><span data-ttu-id="ec7eb-871">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ec7eb-871">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec7eb-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-873">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="ec7eb-873">Office on the web</span></span></td>
    <td> <span data-ttu-id="ec7eb-874">- Контент</span><span class="sxs-lookup"><span data-stu-id="ec7eb-874">- Content</span></span><br><span data-ttu-id="ec7eb-875">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-875">
         - TaskPane</span></span><br><span data-ttu-id="ec7eb-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec7eb-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-880">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec7eb-880">- DocumentEvents</span></span><br><span data-ttu-id="ec7eb-881">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-881">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec7eb-882">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="ec7eb-882">
         - Settings</span></span><br><span data-ttu-id="ec7eb-883">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-883">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="ec7eb-884">Project</span><span class="sxs-lookup"><span data-stu-id="ec7eb-884">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec7eb-885">Платформа</span><span class="sxs-lookup"><span data-stu-id="ec7eb-885">Platform</span></span></th>
    <th><span data-ttu-id="ec7eb-886">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="ec7eb-886">Extension points</span></span></th>
    <th><span data-ttu-id="ec7eb-887">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ec7eb-887">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec7eb-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-889">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-889">Office 2019 on Windows</span></span><br><span data-ttu-id="ec7eb-890">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-890">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-891">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-891">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-893">- Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-893">- Selection</span></span><br><span data-ttu-id="ec7eb-894">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-894">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-895">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-895">Office 2016 on Windows</span></span><br><span data-ttu-id="ec7eb-896">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-896">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-897">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-897">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-899">- Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-899">- Selection</span></span><br><span data-ttu-id="ec7eb-900">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-900">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec7eb-901">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="ec7eb-901">Office 2013 on Windows</span></span><br><span data-ttu-id="ec7eb-902">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-902">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec7eb-903">- Область задач</span><span class="sxs-lookup"><span data-stu-id="ec7eb-903">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec7eb-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec7eb-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec7eb-905">- Selection</span><span class="sxs-lookup"><span data-stu-id="ec7eb-905">- Selection</span></span><br><span data-ttu-id="ec7eb-906">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec7eb-906">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ec7eb-907">См. также</span><span class="sxs-lookup"><span data-stu-id="ec7eb-907">See also</span></span>

- [<span data-ttu-id="ec7eb-908">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ec7eb-908">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ec7eb-909">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="ec7eb-909">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="ec7eb-910">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="ec7eb-910">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="ec7eb-911">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="ec7eb-911">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="ec7eb-912">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="ec7eb-912">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="ec7eb-913">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="ec7eb-913">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="ec7eb-914">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="ec7eb-914">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="ec7eb-915">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="ec7eb-915">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="ec7eb-916">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-916">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="ec7eb-917">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ec7eb-917">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="ec7eb-918">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="ec7eb-918">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
