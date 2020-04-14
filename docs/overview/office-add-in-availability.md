---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 04/13/2020
localization_priority: Priority
ms.openlocfilehash: 72da8db755fe6d1d166f66a70c8c298e5a27adff
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241058"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="6fa89-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6fa89-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="6fa89-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="6fa89-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="6fa89-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="6fa89-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="6fa89-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="6fa89-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="6fa89-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="6fa89-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="6fa89-108">Excel</span><span class="sxs-lookup"><span data-stu-id="6fa89-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6fa89-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="6fa89-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6fa89-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6fa89-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="6fa89-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6fa89-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="6fa89-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6fa89-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6fa89-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="6fa89-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-114">- TaskPane</span></span><br><span data-ttu-id="6fa89-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-115">
        - Content</span></span><br><span data-ttu-id="6fa89-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6fa89-116">
        - Custom Functions</span></span><br><span data-ttu-id="6fa89-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="6fa89-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6fa89-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6fa89-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6fa89-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6fa89-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6fa89-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6fa89-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6fa89-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6fa89-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6fa89-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6fa89-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6fa89-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="6fa89-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6fa89-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-130">
        - BindingEvents</span></span><br><span data-ttu-id="6fa89-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-131">
        - CompressedFile</span></span><br><span data-ttu-id="6fa89-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-132">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-133">
        - File</span></span><br><span data-ttu-id="6fa89-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-134">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-136">
        - Selection</span></span><br><span data-ttu-id="6fa89-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-137">
        - Settings</span></span><br><span data-ttu-id="6fa89-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-138">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-139">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-140">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-142">Office on Windows</span></span><br><span data-ttu-id="6fa89-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-144">- TaskPane</span></span><br><span data-ttu-id="6fa89-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-145">
        - Content</span></span><br><span data-ttu-id="6fa89-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6fa89-146">
        - Custom Functions</span></span><br><span data-ttu-id="6fa89-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="6fa89-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6fa89-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6fa89-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6fa89-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6fa89-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6fa89-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6fa89-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6fa89-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6fa89-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6fa89-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6fa89-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6fa89-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6fa89-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="6fa89-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-161">
        - BindingEvents</span></span><br><span data-ttu-id="6fa89-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-162">
        - CompressedFile</span></span><br><span data-ttu-id="6fa89-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-163">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-164">
        - File</span></span><br><span data-ttu-id="6fa89-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-165">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-167">
        - Selection</span></span><br><span data-ttu-id="6fa89-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-168">
        - Settings</span></span><br><span data-ttu-id="6fa89-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-169">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-170">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-171">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-173">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-173">Office 2019 on Windows</span></span><br><span data-ttu-id="6fa89-174">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6fa89-175">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-175">- TaskPane</span></span><br><span data-ttu-id="6fa89-176">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-176">
        - Content</span></span><br><span data-ttu-id="6fa89-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6fa89-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6fa89-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6fa89-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6fa89-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6fa89-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6fa89-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6fa89-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6fa89-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6fa89-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6fa89-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-188">- BindingEvents</span></span><br><span data-ttu-id="6fa89-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-189">
        - CompressedFile</span></span><br><span data-ttu-id="6fa89-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-190">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-191">
        - File</span></span><br><span data-ttu-id="6fa89-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-192">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-194">
        - Selection</span></span><br><span data-ttu-id="6fa89-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-195">
        - Settings</span></span><br><span data-ttu-id="6fa89-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-196">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-197">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-198">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-200">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-200">Office 2016 on Windows</span></span><br><span data-ttu-id="6fa89-201">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6fa89-202">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-202">- TaskPane</span></span><br><span data-ttu-id="6fa89-203">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-203">
        - Content</span></span></td>
    <td><span data-ttu-id="6fa89-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6fa89-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6fa89-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6fa89-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6fa89-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-207">- BindingEvents</span></span><br><span data-ttu-id="6fa89-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-208">
        - CompressedFile</span></span><br><span data-ttu-id="6fa89-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-209">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-210">
        - File</span></span><br><span data-ttu-id="6fa89-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-211">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-213">
        - Selection</span></span><br><span data-ttu-id="6fa89-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-214">
        - Settings</span></span><br><span data-ttu-id="6fa89-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-215">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-216">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-217">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-219">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-219">Office 2013 on Windows</span></span><br><span data-ttu-id="6fa89-220">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6fa89-221">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-221">
        - TaskPane</span></span><br><span data-ttu-id="6fa89-222">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="6fa89-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6fa89-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6fa89-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6fa89-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-225">
        - BindingEvents</span></span><br><span data-ttu-id="6fa89-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-226">
        - CompressedFile</span></span><br><span data-ttu-id="6fa89-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-227">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-228">
        - File</span></span><br><span data-ttu-id="6fa89-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-229">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-231">
        - Selection</span></span><br><span data-ttu-id="6fa89-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-232">
        - Settings</span></span><br><span data-ttu-id="6fa89-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-233">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-234">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-235">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-237">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="6fa89-237">Office on iPad</span></span><br><span data-ttu-id="6fa89-238">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="6fa89-239">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-239">- TaskPane</span></span><br><span data-ttu-id="6fa89-240">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-240">
        - Content</span></span></td>
    <td><span data-ttu-id="6fa89-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6fa89-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6fa89-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6fa89-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6fa89-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6fa89-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6fa89-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6fa89-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6fa89-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6fa89-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6fa89-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6fa89-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-253">- BindingEvents</span></span><br><span data-ttu-id="6fa89-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-254">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-255">
        - File</span></span><br><span data-ttu-id="6fa89-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-256">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-258">
        - Selection</span></span><br><span data-ttu-id="6fa89-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-259">
        - Settings</span></span><br><span data-ttu-id="6fa89-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-260">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-261">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-262">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-264">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-264">Office on Mac</span></span><br><span data-ttu-id="6fa89-265">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="6fa89-266">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-266">- TaskPane</span></span><br><span data-ttu-id="6fa89-267">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-267">
        - Content</span></span><br><span data-ttu-id="6fa89-268">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6fa89-268">
        - Custom Functions</span></span><br><span data-ttu-id="6fa89-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6fa89-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6fa89-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6fa89-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6fa89-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6fa89-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6fa89-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6fa89-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6fa89-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6fa89-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6fa89-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6fa89-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6fa89-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="6fa89-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-283">- BindingEvents</span></span><br><span data-ttu-id="6fa89-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-284">
        - CompressedFile</span></span><br><span data-ttu-id="6fa89-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-285">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-286">
        - File</span></span><br><span data-ttu-id="6fa89-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-287">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-289">
        - PdfFile</span></span><br><span data-ttu-id="6fa89-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-290">
        - Selection</span></span><br><span data-ttu-id="6fa89-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-291">
        - Settings</span></span><br><span data-ttu-id="6fa89-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-292">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-293">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-294">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-296">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-296">Office 2019 on Mac</span></span><br><span data-ttu-id="6fa89-297">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6fa89-298">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-298">- TaskPane</span></span><br><span data-ttu-id="6fa89-299">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-299">
        - Content</span></span><br><span data-ttu-id="6fa89-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6fa89-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6fa89-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6fa89-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6fa89-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6fa89-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6fa89-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6fa89-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6fa89-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6fa89-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6fa89-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-311">- BindingEvents</span></span><br><span data-ttu-id="6fa89-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-312">
        - CompressedFile</span></span><br><span data-ttu-id="6fa89-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-313">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-314">
        - File</span></span><br><span data-ttu-id="6fa89-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-315">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-317">
        - PdfFile</span></span><br><span data-ttu-id="6fa89-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-318">
        - Selection</span></span><br><span data-ttu-id="6fa89-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-319">
        - Settings</span></span><br><span data-ttu-id="6fa89-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-320">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-321">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-322">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-324">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-324">Office 2016 on Mac</span></span><br><span data-ttu-id="6fa89-325">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6fa89-326">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-326">- TaskPane</span></span><br><span data-ttu-id="6fa89-327">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-327">
        - Content</span></span></td>
    <td><span data-ttu-id="6fa89-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6fa89-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6fa89-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6fa89-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6fa89-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-331">- BindingEvents</span></span><br><span data-ttu-id="6fa89-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-332">
        - CompressedFile</span></span><br><span data-ttu-id="6fa89-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-333">
        - DocumentEvents</span></span><br><span data-ttu-id="6fa89-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-334">
        - File</span></span><br><span data-ttu-id="6fa89-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-335">
        - MatrixBindings</span></span><br><span data-ttu-id="6fa89-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-337">
        - PdfFile</span></span><br><span data-ttu-id="6fa89-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-338">
        - Selection</span></span><br><span data-ttu-id="6fa89-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-339">
        - Settings</span></span><br><span data-ttu-id="6fa89-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-340">
        - TableBindings</span></span><br><span data-ttu-id="6fa89-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-341">
        - TableCoercion</span></span><br><span data-ttu-id="6fa89-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-342">
        - TextBindings</span></span><br><span data-ttu-id="6fa89-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="6fa89-344">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="6fa89-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="6fa89-345">Пользовательские функции (только Excel)</span><span class="sxs-lookup"><span data-stu-id="6fa89-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6fa89-346">Платформа</span><span class="sxs-lookup"><span data-stu-id="6fa89-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6fa89-347">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6fa89-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="6fa89-348">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6fa89-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="6fa89-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6fa89-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-350">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6fa89-350">Office on the web</span></span></td>
    <td><span data-ttu-id="6fa89-351">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6fa89-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="6fa89-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-353">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-353">Office on Windows</span></span><br><span data-ttu-id="6fa89-354">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="6fa89-355">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6fa89-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="6fa89-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-357">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-357">Office for Mac</span></span><br><span data-ttu-id="6fa89-358">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="6fa89-359">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6fa89-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="6fa89-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="6fa89-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="6fa89-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6fa89-362">Платформа</span><span class="sxs-lookup"><span data-stu-id="6fa89-362">Platform</span></span></th>
    <th><span data-ttu-id="6fa89-363">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6fa89-363">Extension points</span></span></th>
    <th><span data-ttu-id="6fa89-364">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6fa89-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="6fa89-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6fa89-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-366">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6fa89-366">Office on the web</span></span><br><span data-ttu-id="6fa89-367">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="6fa89-367">(modern)</span></span></td>
    <td> <span data-ttu-id="6fa89-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6fa89-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6fa89-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6fa89-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6fa89-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="6fa89-381">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-382">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6fa89-382">Office on the web</span></span><br><span data-ttu-id="6fa89-383">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="6fa89-383">(classic)</span></span></td>
    <td> <span data-ttu-id="6fa89-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6fa89-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6fa89-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6fa89-395">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-396">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-396">Office on Windows</span></span><br><span data-ttu-id="6fa89-397">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6fa89-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6fa89-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="6fa89-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6fa89-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6fa89-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6fa89-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="6fa89-412">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-413">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-413">Office 2019 on Windows</span></span><br><span data-ttu-id="6fa89-414">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6fa89-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6fa89-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="6fa89-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6fa89-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6fa89-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="6fa89-428">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-429">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-429">Office 2016 on Windows</span></span><br><span data-ttu-id="6fa89-430">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6fa89-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6fa89-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="6fa89-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="6fa89-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="6fa89-441">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-442">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-442">Office 2013 on Windows</span></span><br><span data-ttu-id="6fa89-443">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="6fa89-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="6fa89-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="6fa89-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="6fa89-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="6fa89-452">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-453">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="6fa89-453">Office on iOS</span></span><br><span data-ttu-id="6fa89-454">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6fa89-462">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-463">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-463">Office on Mac</span></span><br><span data-ttu-id="6fa89-464">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6fa89-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6fa89-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6fa89-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6fa89-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="6fa89-478">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-479">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-479">Office 2019 on Mac</span></span><br><span data-ttu-id="6fa89-480">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6fa89-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6fa89-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6fa89-492">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-493">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-493">Office 2016 on Mac</span></span><br><span data-ttu-id="6fa89-494">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6fa89-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6fa89-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6fa89-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6fa89-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6fa89-506">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-507">Office для Android</span><span class="sxs-lookup"><span data-stu-id="6fa89-507">Office on Android</span></span><br><span data-ttu-id="6fa89-508">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6fa89-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Организатор встречи (создание): собрание по сети</a> (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="6fa89-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="6fa89-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6fa89-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6fa89-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6fa89-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6fa89-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6fa89-517">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6fa89-517">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="6fa89-518">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="6fa89-518">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6fa89-519">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="6fa89-519">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="6fa89-520">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="6fa89-520">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="6fa89-521">Word</span><span class="sxs-lookup"><span data-stu-id="6fa89-521">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6fa89-522">Платформа</span><span class="sxs-lookup"><span data-stu-id="6fa89-522">Platform</span></span></th>
    <th><span data-ttu-id="6fa89-523">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6fa89-523">Extension points</span></span></th>
    <th><span data-ttu-id="6fa89-524">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6fa89-524">API requirement sets</span></span></th>
    <th><span data-ttu-id="6fa89-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6fa89-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-526">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6fa89-526">Office on the web</span></span></td>
    <td> <span data-ttu-id="6fa89-527">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-527">- TaskPane</span></span><br><span data-ttu-id="6fa89-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6fa89-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6fa89-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6fa89-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6fa89-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6fa89-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-535">- BindingEvents</span></span><br><span data-ttu-id="6fa89-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-537">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-538">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-538">
         - File</span></span><br><span data-ttu-id="6fa89-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-540">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-543">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-544">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-544">
         - Selection</span></span><br><span data-ttu-id="6fa89-545">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-545">
         - Settings</span></span><br><span data-ttu-id="6fa89-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-546">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-547">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-548">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-549">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-550">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-551">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-551">Office on Windows</span></span><br><span data-ttu-id="6fa89-552">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-552">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-553">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-553">- TaskPane</span></span><br><span data-ttu-id="6fa89-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6fa89-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6fa89-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6fa89-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6fa89-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6fa89-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-561">- BindingEvents</span></span><br><span data-ttu-id="6fa89-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-562">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-564">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-565">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-565">
         - File</span></span><br><span data-ttu-id="6fa89-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-567">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-570">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-571">
         - Selection</span></span><br><span data-ttu-id="6fa89-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-572">
         - Settings</span></span><br><span data-ttu-id="6fa89-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-573">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-574">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-575">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-576">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-578">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-578">Office 2019 on Windows</span></span><br><span data-ttu-id="6fa89-579">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-580">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-580">- TaskPane</span></span><br><span data-ttu-id="6fa89-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6fa89-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6fa89-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6fa89-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-587">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-587">- BindingEvents</span></span><br><span data-ttu-id="6fa89-588">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-588">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-589">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-589">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-590">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-590">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-591">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-591">
         - File</span></span><br><span data-ttu-id="6fa89-592">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-592">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-593">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-593">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-594">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-594">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-595">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-595">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-596">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-597">
         - Selection</span></span><br><span data-ttu-id="6fa89-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-598">
         - Settings</span></span><br><span data-ttu-id="6fa89-599">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-599">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-600">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-600">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-601">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-601">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-602">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-602">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-603">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-603">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-604">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-604">Office 2016 on Windows</span></span><br><span data-ttu-id="6fa89-605">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-605">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-606">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-606">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6fa89-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6fa89-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6fa89-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-610">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-610">- BindingEvents</span></span><br><span data-ttu-id="6fa89-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-611">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-612">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-612">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-613">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-613">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-614">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-614">
         - File</span></span><br><span data-ttu-id="6fa89-615">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-615">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-616">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-616">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-617">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-617">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-618">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-618">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-619">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-619">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-620">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-620">
         - Selection</span></span><br><span data-ttu-id="6fa89-621">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-621">
         - Settings</span></span><br><span data-ttu-id="6fa89-622">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-622">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-623">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-623">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-624">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-624">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-625">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-626">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-626">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-627">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-627">Office 2013 on Windows</span></span><br><span data-ttu-id="6fa89-628">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-628">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-629">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-629">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6fa89-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6fa89-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-632">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-632">- BindingEvents</span></span><br><span data-ttu-id="6fa89-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-633">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-634">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-634">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-635">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-636">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-636">
         - File</span></span><br><span data-ttu-id="6fa89-637">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-637">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-638">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-638">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-639">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-639">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-640">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-640">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-641">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-641">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-642">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-642">
         - Selection</span></span><br><span data-ttu-id="6fa89-643">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-643">
         - Settings</span></span><br><span data-ttu-id="6fa89-644">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-644">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-645">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-645">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-646">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-646">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-647">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-647">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-648">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-648">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-649">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="6fa89-649">Office on iPad</span></span><br><span data-ttu-id="6fa89-650">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-650">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-651">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6fa89-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6fa89-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6fa89-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="6fa89-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-657">- BindingEvents</span></span><br><span data-ttu-id="6fa89-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-658">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-660">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-661">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-661">
         - File</span></span><br><span data-ttu-id="6fa89-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-663">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-666">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-667">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-667">
         - Selection</span></span><br><span data-ttu-id="6fa89-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-668">
         - Settings</span></span><br><span data-ttu-id="6fa89-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-669">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-670">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-671">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-672">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-674">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-674">Office on Mac</span></span><br><span data-ttu-id="6fa89-675">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-675">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-676">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-676">- TaskPane</span></span><br><span data-ttu-id="6fa89-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6fa89-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6fa89-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6fa89-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6fa89-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="6fa89-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-684">- BindingEvents</span></span><br><span data-ttu-id="6fa89-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-685">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-687">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-688">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-688">
         - File</span></span><br><span data-ttu-id="6fa89-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-690">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-693">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-694">
         - Selection</span></span><br><span data-ttu-id="6fa89-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-695">
         - Settings</span></span><br><span data-ttu-id="6fa89-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-696">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-697">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-698">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-699">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-701">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-701">Office 2019 on Mac</span></span><br><span data-ttu-id="6fa89-702">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-703">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-703">- TaskPane</span></span><br><span data-ttu-id="6fa89-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6fa89-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6fa89-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6fa89-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="6fa89-710">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-710">- BindingEvents</span></span><br><span data-ttu-id="6fa89-711">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-711">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-712">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-712">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-713">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-713">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-714">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-714">
         - File</span></span><br><span data-ttu-id="6fa89-715">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-715">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-716">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-716">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-717">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-717">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-718">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-718">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-719">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-719">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-720">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-720">
         - Selection</span></span><br><span data-ttu-id="6fa89-721">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-721">
         - Settings</span></span><br><span data-ttu-id="6fa89-722">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-722">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-723">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-723">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-724">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-724">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-725">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-726">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-726">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-727">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-727">Office 2016 on Mac</span></span><br><span data-ttu-id="6fa89-728">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-728">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-729">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-729">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6fa89-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6fa89-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6fa89-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-733">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-733">- BindingEvents</span></span><br><span data-ttu-id="6fa89-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-734">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-735">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6fa89-735">
         - CustomXmlParts</span></span><br><span data-ttu-id="6fa89-736">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-736">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-737">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6fa89-737">
         - File</span></span><br><span data-ttu-id="6fa89-738">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-738">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-739">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-739">
         - MatrixBindings</span></span><br><span data-ttu-id="6fa89-740">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-740">
         - MatrixCoercion</span></span><br><span data-ttu-id="6fa89-741">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-741">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6fa89-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-742">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-743">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-743">
         - Selection</span></span><br><span data-ttu-id="6fa89-744">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6fa89-744">
         - Settings</span></span><br><span data-ttu-id="6fa89-745">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-745">
         - TableBindings</span></span><br><span data-ttu-id="6fa89-746">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-746">
         - TableCoercion</span></span><br><span data-ttu-id="6fa89-747">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6fa89-747">
         - TextBindings</span></span><br><span data-ttu-id="6fa89-748">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-748">
         - TextCoercion</span></span><br><span data-ttu-id="6fa89-749">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-749">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="6fa89-750">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="6fa89-750">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="6fa89-751">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6fa89-751">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6fa89-752">Платформа</span><span class="sxs-lookup"><span data-stu-id="6fa89-752">Platform</span></span></th>
    <th><span data-ttu-id="6fa89-753">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6fa89-753">Extension points</span></span></th>
    <th><span data-ttu-id="6fa89-754">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6fa89-754">API requirement sets</span></span></th>
    <th><span data-ttu-id="6fa89-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6fa89-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-756">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6fa89-756">Office on the web</span></span></td>
    <td> <span data-ttu-id="6fa89-757">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-757">- Content</span></span><br><span data-ttu-id="6fa89-758">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-758">
         - TaskPane</span></span><br><span data-ttu-id="6fa89-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6fa89-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6fa89-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6fa89-764">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-764">- ActiveView</span></span><br><span data-ttu-id="6fa89-765">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-765">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-766">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-766">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-767">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-767">
         - File</span></span><br><span data-ttu-id="6fa89-768">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-768">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-769">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-769">
         - Selection</span></span><br><span data-ttu-id="6fa89-770">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-770">
         - Settings</span></span><br><span data-ttu-id="6fa89-771">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-771">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-772">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-772">Office on Windows</span></span><br><span data-ttu-id="6fa89-773">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-773">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-774">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-774">- Content</span></span><br><span data-ttu-id="6fa89-775">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-775">
         - TaskPane</span></span><br><span data-ttu-id="6fa89-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6fa89-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6fa89-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6fa89-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-781">- ActiveView</span></span><br><span data-ttu-id="6fa89-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-782">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-783">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-784">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-784">
         - File</span></span><br><span data-ttu-id="6fa89-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-785">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-786">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-786">
         - Selection</span></span><br><span data-ttu-id="6fa89-787">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-787">
         - Settings</span></span><br><span data-ttu-id="6fa89-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-789">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-789">Office 2019 on Windows</span></span><br><span data-ttu-id="6fa89-790">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-791">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-791">- Content</span></span><br><span data-ttu-id="6fa89-792">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-792">
         - TaskPane</span></span><br><span data-ttu-id="6fa89-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-796">- ActiveView</span></span><br><span data-ttu-id="6fa89-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-797">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-798">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-799">
         - File</span></span><br><span data-ttu-id="6fa89-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-800">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-801">
         - Selection</span></span><br><span data-ttu-id="6fa89-802">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-802">
         - Settings</span></span><br><span data-ttu-id="6fa89-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-804">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-804">Office 2016 on Windows</span></span><br><span data-ttu-id="6fa89-805">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-805">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-806">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-806">- Content</span></span><br><span data-ttu-id="6fa89-807">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6fa89-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6fa89-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-810">- ActiveView</span></span><br><span data-ttu-id="6fa89-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-811">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-812">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-813">
         - File</span></span><br><span data-ttu-id="6fa89-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-814">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-815">
         - Selection</span></span><br><span data-ttu-id="6fa89-816">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-816">
         - Settings</span></span><br><span data-ttu-id="6fa89-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-818">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-818">Office 2013 on Windows</span></span><br><span data-ttu-id="6fa89-819">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-819">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-820">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-820">- Content</span></span><br><span data-ttu-id="6fa89-821">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-821">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="6fa89-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6fa89-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6fa89-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-824">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-824">- ActiveView</span></span><br><span data-ttu-id="6fa89-825">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-825">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-826">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-826">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-827">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-827">
         - File</span></span><br><span data-ttu-id="6fa89-828">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-828">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-829">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-829">
         - Selection</span></span><br><span data-ttu-id="6fa89-830">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-830">
         - Settings</span></span><br><span data-ttu-id="6fa89-831">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-831">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-832">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="6fa89-832">Office on iPad</span></span><br><span data-ttu-id="6fa89-833">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-833">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-834">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-834">- Content</span></span><br><span data-ttu-id="6fa89-835">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-835">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6fa89-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-839">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-839">- ActiveView</span></span><br><span data-ttu-id="6fa89-840">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-840">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-841">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-841">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-842">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-842">
         - File</span></span><br><span data-ttu-id="6fa89-843">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-843">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-844">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-844">
         - Selection</span></span><br><span data-ttu-id="6fa89-845">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-845">
         - Settings</span></span><br><span data-ttu-id="6fa89-846">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-846">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-847">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-847">Office on Mac</span></span><br><span data-ttu-id="6fa89-848">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6fa89-848">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6fa89-849">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-849">- Content</span></span><br><span data-ttu-id="6fa89-850">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-850">
         - TaskPane</span></span><br><span data-ttu-id="6fa89-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6fa89-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6fa89-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6fa89-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-856">- ActiveView</span></span><br><span data-ttu-id="6fa89-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-857">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-858">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-859">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-859">
         - File</span></span><br><span data-ttu-id="6fa89-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-860">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-861">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-861">
         - Selection</span></span><br><span data-ttu-id="6fa89-862">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-862">
         - Settings</span></span><br><span data-ttu-id="6fa89-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-863">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-864">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-864">Office 2019 on Mac</span></span><br><span data-ttu-id="6fa89-865">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-866">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-866">- Content</span></span><br><span data-ttu-id="6fa89-867">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-867">
         - TaskPane</span></span><br><span data-ttu-id="6fa89-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-871">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-871">- ActiveView</span></span><br><span data-ttu-id="6fa89-872">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-872">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-873">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-873">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-874">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-874">
         - File</span></span><br><span data-ttu-id="6fa89-875">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-875">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-876">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-876">
         - Selection</span></span><br><span data-ttu-id="6fa89-877">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-877">
         - Settings</span></span><br><span data-ttu-id="6fa89-878">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-878">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-879">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-879">Office 2016 on Mac</span></span><br><span data-ttu-id="6fa89-880">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-880">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-881">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-881">- Content</span></span><br><span data-ttu-id="6fa89-882">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-882">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6fa89-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6fa89-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-885">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6fa89-885">- ActiveView</span></span><br><span data-ttu-id="6fa89-886">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-886">
         - CompressedFile</span></span><br><span data-ttu-id="6fa89-887">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-887">
         - DocumentEvents</span></span><br><span data-ttu-id="6fa89-888">
         - File</span><span class="sxs-lookup"><span data-stu-id="6fa89-888">
         - File</span></span><br><span data-ttu-id="6fa89-889">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6fa89-889">
         - PdfFile</span></span><br><span data-ttu-id="6fa89-890">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-890">
         - Selection</span></span><br><span data-ttu-id="6fa89-891">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-891">
         - Settings</span></span><br><span data-ttu-id="6fa89-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-892">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="6fa89-893">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="6fa89-893">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="6fa89-894">OneNote</span><span class="sxs-lookup"><span data-stu-id="6fa89-894">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6fa89-895">Платформа</span><span class="sxs-lookup"><span data-stu-id="6fa89-895">Platform</span></span></th>
    <th><span data-ttu-id="6fa89-896">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6fa89-896">Extension points</span></span></th>
    <th><span data-ttu-id="6fa89-897">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6fa89-897">API requirement sets</span></span></th>
    <th><span data-ttu-id="6fa89-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6fa89-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-899">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6fa89-899">Office on the web</span></span></td>
    <td> <span data-ttu-id="6fa89-900">- Контент</span><span class="sxs-lookup"><span data-stu-id="6fa89-900">- Content</span></span><br><span data-ttu-id="6fa89-901">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-901">
         - TaskPane</span></span><br><span data-ttu-id="6fa89-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6fa89-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="6fa89-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6fa89-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-906">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6fa89-906">- DocumentEvents</span></span><br><span data-ttu-id="6fa89-907">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-907">
         - HtmlCoercion</span></span><br><span data-ttu-id="6fa89-908">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6fa89-908">
         - Settings</span></span><br><span data-ttu-id="6fa89-909">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-909">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="6fa89-910">Project</span><span class="sxs-lookup"><span data-stu-id="6fa89-910">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6fa89-911">Платформа</span><span class="sxs-lookup"><span data-stu-id="6fa89-911">Platform</span></span></th>
    <th><span data-ttu-id="6fa89-912">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6fa89-912">Extension points</span></span></th>
    <th><span data-ttu-id="6fa89-913">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6fa89-913">API requirement sets</span></span></th>
    <th><span data-ttu-id="6fa89-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6fa89-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-915">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-915">Office 2019 on Windows</span></span><br><span data-ttu-id="6fa89-916">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-916">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-917">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-917">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-919">- Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-919">- Selection</span></span><br><span data-ttu-id="6fa89-920">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-920">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-921">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-921">Office 2016 on Windows</span></span><br><span data-ttu-id="6fa89-922">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-922">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-923">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-923">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-925">- Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-925">- Selection</span></span><br><span data-ttu-id="6fa89-926">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-926">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6fa89-927">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6fa89-927">Office 2013 on Windows</span></span><br><span data-ttu-id="6fa89-928">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6fa89-928">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6fa89-929">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6fa89-929">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6fa89-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6fa89-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6fa89-931">- Selection</span><span class="sxs-lookup"><span data-stu-id="6fa89-931">- Selection</span></span><br><span data-ttu-id="6fa89-932">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6fa89-932">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="6fa89-933">См. также</span><span class="sxs-lookup"><span data-stu-id="6fa89-933">See also</span></span>

- [<span data-ttu-id="6fa89-934">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6fa89-934">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="6fa89-935">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="6fa89-935">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="6fa89-936">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="6fa89-936">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="6fa89-937">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="6fa89-937">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="6fa89-938">Справочная документация по API</span><span class="sxs-lookup"><span data-stu-id="6fa89-938">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="6fa89-939">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="6fa89-939">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="6fa89-940">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="6fa89-940">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="6fa89-941">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="6fa89-941">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="6fa89-942">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="6fa89-942">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="6fa89-943">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="6fa89-943">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="6fa89-944">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6fa89-944">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="6fa89-945">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6fa89-945">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)