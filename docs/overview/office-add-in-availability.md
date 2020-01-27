---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 01/23/2020
localization_priority: Priority
ms.openlocfilehash: b30fe872fd89bb02afac99a7838d43d1fbee5464
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554022"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="9de31-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9de31-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="9de31-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="9de31-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="9de31-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="9de31-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="9de31-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="9de31-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="9de31-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="9de31-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="9de31-108">Excel</span><span class="sxs-lookup"><span data-stu-id="9de31-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9de31-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="9de31-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9de31-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9de31-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9de31-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9de31-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9de31-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9de31-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9de31-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="9de31-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-114">- TaskPane</span></span><br><span data-ttu-id="9de31-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-115">
        - Content</span></span><br><span data-ttu-id="9de31-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9de31-116">
        - Custom Functions</span></span><br><span data-ttu-id="9de31-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="9de31-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9de31-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9de31-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9de31-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9de31-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9de31-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9de31-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9de31-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9de31-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9de31-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="9de31-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="9de31-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="9de31-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="9de31-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="9de31-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="9de31-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9de31-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-130">
        - BindingEvents</span></span><br><span data-ttu-id="9de31-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-131">
        - CompressedFile</span></span><br><span data-ttu-id="9de31-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-132">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-133">
        - File</span></span><br><span data-ttu-id="9de31-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-134">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-136">
        - Selection</span></span><br><span data-ttu-id="9de31-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-137">
        - Settings</span></span><br><span data-ttu-id="9de31-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-138">
        - TableBindings</span></span><br><span data-ttu-id="9de31-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-139">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-140">
        - TextBindings</span></span><br><span data-ttu-id="9de31-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-142">Office on Windows</span></span><br><span data-ttu-id="9de31-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-144">- TaskPane</span></span><br><span data-ttu-id="9de31-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-145">
        - Content</span></span><br><span data-ttu-id="9de31-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9de31-146">
        - Custom Functions</span></span><br><span data-ttu-id="9de31-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="9de31-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9de31-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9de31-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9de31-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9de31-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9de31-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9de31-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9de31-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9de31-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9de31-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="9de31-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="9de31-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="9de31-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="9de31-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9de31-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="9de31-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-161">
        - BindingEvents</span></span><br><span data-ttu-id="9de31-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-162">
        - CompressedFile</span></span><br><span data-ttu-id="9de31-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-163">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-164">
        - File</span></span><br><span data-ttu-id="9de31-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-165">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-167">
        - Selection</span></span><br><span data-ttu-id="9de31-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-168">
        - Settings</span></span><br><span data-ttu-id="9de31-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-169">
        - TableBindings</span></span><br><span data-ttu-id="9de31-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-170">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-171">
        - TextBindings</span></span><br><span data-ttu-id="9de31-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-173">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-173">Office 2019 on Windows</span></span><br><span data-ttu-id="9de31-174">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9de31-175">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-175">- TaskPane</span></span><br><span data-ttu-id="9de31-176">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-176">
        - Content</span></span><br><span data-ttu-id="9de31-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9de31-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9de31-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9de31-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9de31-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9de31-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9de31-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9de31-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9de31-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9de31-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9de31-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-188">- BindingEvents</span></span><br><span data-ttu-id="9de31-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-189">
        - CompressedFile</span></span><br><span data-ttu-id="9de31-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-190">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-191">
        - File</span></span><br><span data-ttu-id="9de31-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-192">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-194">
        - Selection</span></span><br><span data-ttu-id="9de31-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-195">
        - Settings</span></span><br><span data-ttu-id="9de31-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-196">
        - TableBindings</span></span><br><span data-ttu-id="9de31-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-197">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-198">
        - TextBindings</span></span><br><span data-ttu-id="9de31-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-200">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-200">Office 2016 on Windows</span></span><br><span data-ttu-id="9de31-201">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9de31-202">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-202">- TaskPane</span></span><br><span data-ttu-id="9de31-203">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-203">
        - Content</span></span></td>
    <td><span data-ttu-id="9de31-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9de31-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9de31-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="9de31-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9de31-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-207">- BindingEvents</span></span><br><span data-ttu-id="9de31-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-208">
        - CompressedFile</span></span><br><span data-ttu-id="9de31-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-209">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-210">
        - File</span></span><br><span data-ttu-id="9de31-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-211">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-213">
        - Selection</span></span><br><span data-ttu-id="9de31-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-214">
        - Settings</span></span><br><span data-ttu-id="9de31-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-215">
        - TableBindings</span></span><br><span data-ttu-id="9de31-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-216">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-217">
        - TextBindings</span></span><br><span data-ttu-id="9de31-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-219">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-219">Office 2013 on Windows</span></span><br><span data-ttu-id="9de31-220">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9de31-221">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-221">
        - TaskPane</span></span><br><span data-ttu-id="9de31-222">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="9de31-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9de31-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9de31-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9de31-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-225">
        - BindingEvents</span></span><br><span data-ttu-id="9de31-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-226">
        - CompressedFile</span></span><br><span data-ttu-id="9de31-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-227">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-228">
        - File</span></span><br><span data-ttu-id="9de31-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-229">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-231">
        - Selection</span></span><br><span data-ttu-id="9de31-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-232">
        - Settings</span></span><br><span data-ttu-id="9de31-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-233">
        - TableBindings</span></span><br><span data-ttu-id="9de31-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-234">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-235">
        - TextBindings</span></span><br><span data-ttu-id="9de31-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-237">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9de31-237">Office on iPad</span></span><br><span data-ttu-id="9de31-238">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="9de31-239">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-239">- TaskPane</span></span><br><span data-ttu-id="9de31-240">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-240">
        - Content</span></span></td>
    <td><span data-ttu-id="9de31-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9de31-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9de31-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9de31-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9de31-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9de31-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9de31-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9de31-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9de31-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="9de31-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="9de31-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="9de31-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="9de31-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9de31-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-253">- BindingEvents</span></span><br><span data-ttu-id="9de31-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-254">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-255">
        - File</span></span><br><span data-ttu-id="9de31-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-256">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-258">
        - Selection</span></span><br><span data-ttu-id="9de31-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-259">
        - Settings</span></span><br><span data-ttu-id="9de31-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-260">
        - TableBindings</span></span><br><span data-ttu-id="9de31-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-261">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-262">
        - TextBindings</span></span><br><span data-ttu-id="9de31-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-264">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-264">Office on Mac</span></span><br><span data-ttu-id="9de31-265">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="9de31-266">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-266">- TaskPane</span></span><br><span data-ttu-id="9de31-267">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-267">
        - Content</span></span><br><span data-ttu-id="9de31-268">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9de31-268">
        - Custom Functions</span></span><br><span data-ttu-id="9de31-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9de31-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9de31-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9de31-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9de31-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9de31-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9de31-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9de31-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9de31-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9de31-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="9de31-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="9de31-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="9de31-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="9de31-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9de31-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="9de31-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-283">- BindingEvents</span></span><br><span data-ttu-id="9de31-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-284">
        - CompressedFile</span></span><br><span data-ttu-id="9de31-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-285">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-286">
        - File</span></span><br><span data-ttu-id="9de31-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-287">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-289">
        - PdfFile</span></span><br><span data-ttu-id="9de31-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-290">
        - Selection</span></span><br><span data-ttu-id="9de31-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-291">
        - Settings</span></span><br><span data-ttu-id="9de31-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-292">
        - TableBindings</span></span><br><span data-ttu-id="9de31-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-293">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-294">
        - TextBindings</span></span><br><span data-ttu-id="9de31-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-296">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-296">Office 2019 on Mac</span></span><br><span data-ttu-id="9de31-297">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9de31-298">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-298">- TaskPane</span></span><br><span data-ttu-id="9de31-299">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-299">
        - Content</span></span><br><span data-ttu-id="9de31-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9de31-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9de31-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9de31-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9de31-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9de31-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9de31-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9de31-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9de31-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9de31-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9de31-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-311">- BindingEvents</span></span><br><span data-ttu-id="9de31-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-312">
        - CompressedFile</span></span><br><span data-ttu-id="9de31-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-313">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-314">
        - File</span></span><br><span data-ttu-id="9de31-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-315">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-317">
        - PdfFile</span></span><br><span data-ttu-id="9de31-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-318">
        - Selection</span></span><br><span data-ttu-id="9de31-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-319">
        - Settings</span></span><br><span data-ttu-id="9de31-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-320">
        - TableBindings</span></span><br><span data-ttu-id="9de31-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-321">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-322">
        - TextBindings</span></span><br><span data-ttu-id="9de31-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-324">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-324">Office 2016 on Mac</span></span><br><span data-ttu-id="9de31-325">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9de31-326">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-326">- TaskPane</span></span><br><span data-ttu-id="9de31-327">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-327">
        - Content</span></span></td>
    <td><span data-ttu-id="9de31-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9de31-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9de31-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="9de31-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9de31-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-331">- BindingEvents</span></span><br><span data-ttu-id="9de31-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-332">
        - CompressedFile</span></span><br><span data-ttu-id="9de31-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-333">
        - DocumentEvents</span></span><br><span data-ttu-id="9de31-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="9de31-334">
        - File</span></span><br><span data-ttu-id="9de31-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-335">
        - MatrixBindings</span></span><br><span data-ttu-id="9de31-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="9de31-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-337">
        - PdfFile</span></span><br><span data-ttu-id="9de31-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-338">
        - Selection</span></span><br><span data-ttu-id="9de31-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-339">
        - Settings</span></span><br><span data-ttu-id="9de31-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-340">
        - TableBindings</span></span><br><span data-ttu-id="9de31-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-341">
        - TableCoercion</span></span><br><span data-ttu-id="9de31-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-342">
        - TextBindings</span></span><br><span data-ttu-id="9de31-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="9de31-344">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="9de31-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="9de31-345">Пользовательские функции (только Excel)</span><span class="sxs-lookup"><span data-stu-id="9de31-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9de31-346">Платформа</span><span class="sxs-lookup"><span data-stu-id="9de31-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9de31-347">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9de31-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9de31-348">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9de31-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9de31-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9de31-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-350">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9de31-350">Office on the web</span></span></td>
    <td><span data-ttu-id="9de31-351">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9de31-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="9de31-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-353">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-353">Office on Windows</span></span><br><span data-ttu-id="9de31-354">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="9de31-355">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9de31-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="9de31-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-357">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-357">Office for Mac</span></span><br><span data-ttu-id="9de31-358">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="9de31-359">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9de31-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="9de31-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="9de31-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="9de31-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9de31-362">Платформа</span><span class="sxs-lookup"><span data-stu-id="9de31-362">Platform</span></span></th>
    <th><span data-ttu-id="9de31-363">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9de31-363">Extension points</span></span></th>
    <th><span data-ttu-id="9de31-364">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9de31-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="9de31-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9de31-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-366">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9de31-366">Office on the web</span></span><br><span data-ttu-id="9de31-367">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="9de31-367">(modern)</span></span></td>
    <td> <span data-ttu-id="9de31-368">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-368">- Mail Read</span></span><br><span data-ttu-id="9de31-369">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-369">
      - Mail Compose</span></span><br><span data-ttu-id="9de31-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9de31-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9de31-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="9de31-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="9de31-379">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-380">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9de31-380">Office on the web</span></span><br><span data-ttu-id="9de31-381">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="9de31-381">(classic)</span></span></td>
    <td> <span data-ttu-id="9de31-382">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-382">- Mail Read</span></span><br><span data-ttu-id="9de31-383">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-383">
      - Mail Compose</span></span><br><span data-ttu-id="9de31-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9de31-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9de31-391">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-392">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-392">Office on Windows</span></span><br><span data-ttu-id="9de31-393">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-394">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-394">- Mail Read</span></span><br><span data-ttu-id="9de31-395">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-395">
      - Mail Compose</span></span><br><span data-ttu-id="9de31-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9de31-397">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="9de31-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9de31-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9de31-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9de31-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="9de31-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="9de31-406">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-407">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-407">Office 2019 on Windows</span></span><br><span data-ttu-id="9de31-408">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-409">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-409">- Mail Read</span></span><br><span data-ttu-id="9de31-410">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-410">
      - Mail Compose</span></span><br><span data-ttu-id="9de31-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9de31-412">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="9de31-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9de31-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9de31-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9de31-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9de31-420">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-421">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-421">Office 2016 on Windows</span></span><br><span data-ttu-id="9de31-422">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-423">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-423">- Mail Read</span></span><br><span data-ttu-id="9de31-424">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-424">
      - Mail Compose</span></span><br><span data-ttu-id="9de31-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9de31-426">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="9de31-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9de31-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="9de31-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="9de31-431">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-432">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-432">Office 2013 on Windows</span></span><br><span data-ttu-id="9de31-433">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-434">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-434">- Mail Read</span></span><br><span data-ttu-id="9de31-435">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="9de31-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="9de31-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="9de31-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="9de31-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="9de31-440">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-441">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="9de31-441">Office on iOS</span></span><br><span data-ttu-id="9de31-442">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-443">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-443">- Mail Read</span></span><br><span data-ttu-id="9de31-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9de31-450">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-451">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-451">Office on Mac</span></span><br><span data-ttu-id="9de31-452">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-453">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-453">- Mail Read</span></span><br><span data-ttu-id="9de31-454">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-454">
      - Mail Compose</span></span><br><span data-ttu-id="9de31-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9de31-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9de31-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9de31-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="9de31-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9de31-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="9de31-464">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-465">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-465">Office 2019 on Mac</span></span><br><span data-ttu-id="9de31-466">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-467">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-467">- Mail Read</span></span><br><span data-ttu-id="9de31-468">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-468">
      - Mail Compose</span></span><br><span data-ttu-id="9de31-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9de31-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9de31-476">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-477">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-477">Office 2016 on Mac</span></span><br><span data-ttu-id="9de31-478">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-479">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-479">- Mail Read</span></span><br><span data-ttu-id="9de31-480">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9de31-480">
      - Mail Compose</span></span><br><span data-ttu-id="9de31-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9de31-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9de31-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9de31-488">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-489">Office для Android</span><span class="sxs-lookup"><span data-stu-id="9de31-489">Office on Android</span></span><br><span data-ttu-id="9de31-490">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-491">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9de31-491">- Mail Read</span></span><br><span data-ttu-id="9de31-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9de31-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9de31-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9de31-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9de31-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9de31-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9de31-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9de31-498">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9de31-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="9de31-499">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="9de31-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9de31-500">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="9de31-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="9de31-501">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="9de31-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="9de31-502">Word</span><span class="sxs-lookup"><span data-stu-id="9de31-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9de31-503">Платформа</span><span class="sxs-lookup"><span data-stu-id="9de31-503">Platform</span></span></th>
    <th><span data-ttu-id="9de31-504">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9de31-504">Extension points</span></span></th>
    <th><span data-ttu-id="9de31-505">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9de31-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="9de31-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9de31-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-507">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9de31-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="9de31-508">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-508">- TaskPane</span></span><br><span data-ttu-id="9de31-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9de31-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9de31-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9de31-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9de31-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9de31-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-516">- BindingEvents</span></span><br><span data-ttu-id="9de31-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-518">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-519">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-519">
         - File</span></span><br><span data-ttu-id="9de31-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-521">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-524">
         - PdfFile</span></span><br><span data-ttu-id="9de31-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-525">
         - Selection</span></span><br><span data-ttu-id="9de31-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-526">
         - Settings</span></span><br><span data-ttu-id="9de31-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-527">
         - TableBindings</span></span><br><span data-ttu-id="9de31-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-528">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-529">
         - TextBindings</span></span><br><span data-ttu-id="9de31-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-530">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-532">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-532">Office on Windows</span></span><br><span data-ttu-id="9de31-533">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-534">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-534">- TaskPane</span></span><br><span data-ttu-id="9de31-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9de31-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9de31-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9de31-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9de31-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9de31-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-542">- BindingEvents</span></span><br><span data-ttu-id="9de31-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-543">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-545">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-546">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-546">
         - File</span></span><br><span data-ttu-id="9de31-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-548">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-551">
         - PdfFile</span></span><br><span data-ttu-id="9de31-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-552">
         - Selection</span></span><br><span data-ttu-id="9de31-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-553">
         - Settings</span></span><br><span data-ttu-id="9de31-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-554">
         - TableBindings</span></span><br><span data-ttu-id="9de31-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-555">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-556">
         - TextBindings</span></span><br><span data-ttu-id="9de31-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-557">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-559">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-559">Office 2019 on Windows</span></span><br><span data-ttu-id="9de31-560">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-561">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-561">- TaskPane</span></span><br><span data-ttu-id="9de31-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9de31-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9de31-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9de31-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-568">- BindingEvents</span></span><br><span data-ttu-id="9de31-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-569">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-571">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-572">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-572">
         - File</span></span><br><span data-ttu-id="9de31-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-574">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-577">
         - PdfFile</span></span><br><span data-ttu-id="9de31-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-578">
         - Selection</span></span><br><span data-ttu-id="9de31-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-579">
         - Settings</span></span><br><span data-ttu-id="9de31-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-580">
         - TableBindings</span></span><br><span data-ttu-id="9de31-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-581">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-582">
         - TextBindings</span></span><br><span data-ttu-id="9de31-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-583">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-585">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-585">Office 2016 on Windows</span></span><br><span data-ttu-id="9de31-586">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-587">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9de31-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9de31-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="9de31-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-591">- BindingEvents</span></span><br><span data-ttu-id="9de31-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-592">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-594">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-595">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-595">
         - File</span></span><br><span data-ttu-id="9de31-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-597">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-600">
         - PdfFile</span></span><br><span data-ttu-id="9de31-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-601">
         - Selection</span></span><br><span data-ttu-id="9de31-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-602">
         - Settings</span></span><br><span data-ttu-id="9de31-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-603">
         - TableBindings</span></span><br><span data-ttu-id="9de31-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-604">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-605">
         - TextBindings</span></span><br><span data-ttu-id="9de31-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-606">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-608">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-608">Office 2013 on Windows</span></span><br><span data-ttu-id="9de31-609">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-610">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9de31-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9de31-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-613">- BindingEvents</span></span><br><span data-ttu-id="9de31-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-614">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-616">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-617">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-617">
         - File</span></span><br><span data-ttu-id="9de31-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-619">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-622">
         - PdfFile</span></span><br><span data-ttu-id="9de31-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-623">
         - Selection</span></span><br><span data-ttu-id="9de31-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-624">
         - Settings</span></span><br><span data-ttu-id="9de31-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-625">
         - TableBindings</span></span><br><span data-ttu-id="9de31-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-626">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-627">
         - TextBindings</span></span><br><span data-ttu-id="9de31-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-628">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-630">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9de31-630">Office on iPad</span></span><br><span data-ttu-id="9de31-631">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-632">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9de31-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9de31-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9de31-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="9de31-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-638">- BindingEvents</span></span><br><span data-ttu-id="9de31-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-639">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-641">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-642">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-642">
         - File</span></span><br><span data-ttu-id="9de31-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-644">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-647">
         - PdfFile</span></span><br><span data-ttu-id="9de31-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-648">
         - Selection</span></span><br><span data-ttu-id="9de31-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-649">
         - Settings</span></span><br><span data-ttu-id="9de31-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-650">
         - TableBindings</span></span><br><span data-ttu-id="9de31-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-651">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-652">
         - TextBindings</span></span><br><span data-ttu-id="9de31-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-653">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-655">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-655">Office on Mac</span></span><br><span data-ttu-id="9de31-656">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-657">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-657">- TaskPane</span></span><br><span data-ttu-id="9de31-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9de31-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9de31-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9de31-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9de31-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="9de31-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-665">- BindingEvents</span></span><br><span data-ttu-id="9de31-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-666">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-668">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-669">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-669">
         - File</span></span><br><span data-ttu-id="9de31-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-671">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-674">
         - PdfFile</span></span><br><span data-ttu-id="9de31-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-675">
         - Selection</span></span><br><span data-ttu-id="9de31-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-676">
         - Settings</span></span><br><span data-ttu-id="9de31-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-677">
         - TableBindings</span></span><br><span data-ttu-id="9de31-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-678">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-679">
         - TextBindings</span></span><br><span data-ttu-id="9de31-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-680">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-682">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-682">Office 2019 on Mac</span></span><br><span data-ttu-id="9de31-683">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-684">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-684">- TaskPane</span></span><br><span data-ttu-id="9de31-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9de31-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9de31-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9de31-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9de31-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="9de31-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-691">- BindingEvents</span></span><br><span data-ttu-id="9de31-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-692">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-694">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-695">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-695">
         - File</span></span><br><span data-ttu-id="9de31-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-697">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-700">
         - PdfFile</span></span><br><span data-ttu-id="9de31-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-701">
         - Selection</span></span><br><span data-ttu-id="9de31-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-702">
         - Settings</span></span><br><span data-ttu-id="9de31-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-703">
         - TableBindings</span></span><br><span data-ttu-id="9de31-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-704">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-705">
         - TextBindings</span></span><br><span data-ttu-id="9de31-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-706">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-708">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-708">Office 2016 on Mac</span></span><br><span data-ttu-id="9de31-709">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-710">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9de31-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9de31-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="9de31-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-714">- BindingEvents</span></span><br><span data-ttu-id="9de31-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-715">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9de31-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="9de31-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-717">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-718">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9de31-718">
         - File</span></span><br><span data-ttu-id="9de31-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-720">
         - MatrixBindings</span></span><br><span data-ttu-id="9de31-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="9de31-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9de31-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-723">
         - PdfFile</span></span><br><span data-ttu-id="9de31-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-724">
         - Selection</span></span><br><span data-ttu-id="9de31-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9de31-725">
         - Settings</span></span><br><span data-ttu-id="9de31-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-726">
         - TableBindings</span></span><br><span data-ttu-id="9de31-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-727">
         - TableCoercion</span></span><br><span data-ttu-id="9de31-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9de31-728">
         - TextBindings</span></span><br><span data-ttu-id="9de31-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-729">
         - TextCoercion</span></span><br><span data-ttu-id="9de31-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9de31-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="9de31-731">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="9de31-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="9de31-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9de31-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9de31-733">Платформа</span><span class="sxs-lookup"><span data-stu-id="9de31-733">Platform</span></span></th>
    <th><span data-ttu-id="9de31-734">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9de31-734">Extension points</span></span></th>
    <th><span data-ttu-id="9de31-735">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9de31-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="9de31-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9de31-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-737">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9de31-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="9de31-738">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-738">- Content</span></span><br><span data-ttu-id="9de31-739">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-739">
         - TaskPane</span></span><br><span data-ttu-id="9de31-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="9de31-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9de31-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9de31-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-745">- ActiveView</span></span><br><span data-ttu-id="9de31-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-746">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-747">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-748">
         - File</span></span><br><span data-ttu-id="9de31-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-749">
         - PdfFile</span></span><br><span data-ttu-id="9de31-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-750">
         - Selection</span></span><br><span data-ttu-id="9de31-751">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-751">
         - Settings</span></span><br><span data-ttu-id="9de31-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-753">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-753">Office on Windows</span></span><br><span data-ttu-id="9de31-754">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-755">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-755">- Content</span></span><br><span data-ttu-id="9de31-756">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-756">
         - TaskPane</span></span><br><span data-ttu-id="9de31-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="9de31-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9de31-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9de31-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-762">- ActiveView</span></span><br><span data-ttu-id="9de31-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-763">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-764">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-765">
         - File</span></span><br><span data-ttu-id="9de31-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-766">
         - PdfFile</span></span><br><span data-ttu-id="9de31-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-767">
         - Selection</span></span><br><span data-ttu-id="9de31-768">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-768">
         - Settings</span></span><br><span data-ttu-id="9de31-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-770">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-770">Office 2019 on Windows</span></span><br><span data-ttu-id="9de31-771">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-772">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-772">- Content</span></span><br><span data-ttu-id="9de31-773">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-773">
         - TaskPane</span></span><br><span data-ttu-id="9de31-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-777">- ActiveView</span></span><br><span data-ttu-id="9de31-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-778">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-779">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-780">
         - File</span></span><br><span data-ttu-id="9de31-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-781">
         - PdfFile</span></span><br><span data-ttu-id="9de31-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-782">
         - Selection</span></span><br><span data-ttu-id="9de31-783">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-783">
         - Settings</span></span><br><span data-ttu-id="9de31-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-785">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-785">Office 2016 on Windows</span></span><br><span data-ttu-id="9de31-786">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-787">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-787">- Content</span></span><br><span data-ttu-id="9de31-788">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9de31-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9de31-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-791">- ActiveView</span></span><br><span data-ttu-id="9de31-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-792">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-793">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-794">
         - File</span></span><br><span data-ttu-id="9de31-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-795">
         - PdfFile</span></span><br><span data-ttu-id="9de31-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-796">
         - Selection</span></span><br><span data-ttu-id="9de31-797">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-797">
         - Settings</span></span><br><span data-ttu-id="9de31-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-799">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-799">Office 2013 on Windows</span></span><br><span data-ttu-id="9de31-800">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-801">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-801">- Content</span></span><br><span data-ttu-id="9de31-802">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="9de31-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9de31-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9de31-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-805">- ActiveView</span></span><br><span data-ttu-id="9de31-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-806">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-807">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-808">
         - File</span></span><br><span data-ttu-id="9de31-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-809">
         - PdfFile</span></span><br><span data-ttu-id="9de31-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-810">
         - Selection</span></span><br><span data-ttu-id="9de31-811">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-811">
         - Settings</span></span><br><span data-ttu-id="9de31-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-813">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9de31-813">Office on iPad</span></span><br><span data-ttu-id="9de31-814">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-815">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-815">- Content</span></span><br><span data-ttu-id="9de31-816">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="9de31-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-820">- ActiveView</span></span><br><span data-ttu-id="9de31-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-821">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-822">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-823">
         - File</span></span><br><span data-ttu-id="9de31-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-824">
         - PdfFile</span></span><br><span data-ttu-id="9de31-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-825">
         - Selection</span></span><br><span data-ttu-id="9de31-826">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-826">
         - Settings</span></span><br><span data-ttu-id="9de31-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-828">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-828">Office on Mac</span></span><br><span data-ttu-id="9de31-829">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9de31-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9de31-830">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-830">- Content</span></span><br><span data-ttu-id="9de31-831">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-831">
         - TaskPane</span></span><br><span data-ttu-id="9de31-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="9de31-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9de31-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9de31-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9de31-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-837">- ActiveView</span></span><br><span data-ttu-id="9de31-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-838">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-839">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-840">
         - File</span></span><br><span data-ttu-id="9de31-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-841">
         - PdfFile</span></span><br><span data-ttu-id="9de31-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-842">
         - Selection</span></span><br><span data-ttu-id="9de31-843">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-843">
         - Settings</span></span><br><span data-ttu-id="9de31-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-845">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-845">Office 2019 on Mac</span></span><br><span data-ttu-id="9de31-846">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-847">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-847">- Content</span></span><br><span data-ttu-id="9de31-848">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-848">
         - TaskPane</span></span><br><span data-ttu-id="9de31-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-852">- ActiveView</span></span><br><span data-ttu-id="9de31-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-853">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-854">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-855">
         - File</span></span><br><span data-ttu-id="9de31-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-856">
         - PdfFile</span></span><br><span data-ttu-id="9de31-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-857">
         - Selection</span></span><br><span data-ttu-id="9de31-858">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-858">
         - Settings</span></span><br><span data-ttu-id="9de31-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-860">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-860">Office 2016 on Mac</span></span><br><span data-ttu-id="9de31-861">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-862">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-862">- Content</span></span><br><span data-ttu-id="9de31-863">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9de31-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9de31-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9de31-866">- ActiveView</span></span><br><span data-ttu-id="9de31-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9de31-867">
         - CompressedFile</span></span><br><span data-ttu-id="9de31-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-868">
         - DocumentEvents</span></span><br><span data-ttu-id="9de31-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="9de31-869">
         - File</span></span><br><span data-ttu-id="9de31-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9de31-870">
         - PdfFile</span></span><br><span data-ttu-id="9de31-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-871">
         - Selection</span></span><br><span data-ttu-id="9de31-872">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-872">
         - Settings</span></span><br><span data-ttu-id="9de31-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="9de31-874">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="9de31-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="9de31-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="9de31-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9de31-876">Платформа</span><span class="sxs-lookup"><span data-stu-id="9de31-876">Platform</span></span></th>
    <th><span data-ttu-id="9de31-877">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9de31-877">Extension points</span></span></th>
    <th><span data-ttu-id="9de31-878">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9de31-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="9de31-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9de31-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-880">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9de31-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="9de31-881">- Контент</span><span class="sxs-lookup"><span data-stu-id="9de31-881">- Content</span></span><br><span data-ttu-id="9de31-882">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-882">
         - TaskPane</span></span><br><span data-ttu-id="9de31-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9de31-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9de31-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="9de31-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9de31-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9de31-887">- DocumentEvents</span></span><br><span data-ttu-id="9de31-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="9de31-889">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9de31-889">
         - Settings</span></span><br><span data-ttu-id="9de31-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="9de31-891">Project</span><span class="sxs-lookup"><span data-stu-id="9de31-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9de31-892">Платформа</span><span class="sxs-lookup"><span data-stu-id="9de31-892">Platform</span></span></th>
    <th><span data-ttu-id="9de31-893">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9de31-893">Extension points</span></span></th>
    <th><span data-ttu-id="9de31-894">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9de31-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="9de31-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9de31-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-896">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-896">Office 2019 on Windows</span></span><br><span data-ttu-id="9de31-897">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-898">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-900">- Selection</span></span><br><span data-ttu-id="9de31-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-902">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-902">Office 2016 on Windows</span></span><br><span data-ttu-id="9de31-903">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-904">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-906">- Selection</span></span><br><span data-ttu-id="9de31-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9de31-908">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9de31-908">Office 2013 on Windows</span></span><br><span data-ttu-id="9de31-909">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9de31-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9de31-910">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9de31-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9de31-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9de31-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9de31-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="9de31-912">- Selection</span></span><br><span data-ttu-id="9de31-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9de31-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="9de31-914">См. также</span><span class="sxs-lookup"><span data-stu-id="9de31-914">See also</span></span>

- [<span data-ttu-id="9de31-915">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9de31-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="9de31-916">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="9de31-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="9de31-917">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="9de31-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="9de31-918">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="9de31-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="9de31-919">Справочная документация по API</span><span class="sxs-lookup"><span data-stu-id="9de31-919">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="9de31-920">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="9de31-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="9de31-921">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="9de31-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="9de31-922">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="9de31-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="9de31-923">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="9de31-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="9de31-924">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="9de31-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="9de31-925">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9de31-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="9de31-926">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9de31-926">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)