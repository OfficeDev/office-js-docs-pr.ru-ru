---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: 956ee6b8a9e990a3d6d942ee4a65a1e9275ea025
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851371"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d64ee-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d64ee-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d64ee-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="d64ee-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="d64ee-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="d64ee-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="d64ee-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="d64ee-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="d64ee-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="d64ee-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="d64ee-108">Excel</span><span class="sxs-lookup"><span data-stu-id="d64ee-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d64ee-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="d64ee-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d64ee-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d64ee-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d64ee-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d64ee-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d64ee-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d64ee-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d64ee-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="d64ee-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-114">- TaskPane</span></span><br><span data-ttu-id="d64ee-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-115">
        - Content</span></span><br><span data-ttu-id="d64ee-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d64ee-116">
        - Custom Functions</span></span><br><span data-ttu-id="d64ee-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="d64ee-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d64ee-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d64ee-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d64ee-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d64ee-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d64ee-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d64ee-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d64ee-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d64ee-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d64ee-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d64ee-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d64ee-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="d64ee-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d64ee-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-130">
        - BindingEvents</span></span><br><span data-ttu-id="d64ee-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-131">
        - CompressedFile</span></span><br><span data-ttu-id="d64ee-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-132">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-133">
        - File</span></span><br><span data-ttu-id="d64ee-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-134">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-136">
        - Selection</span></span><br><span data-ttu-id="d64ee-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-137">
        - Settings</span></span><br><span data-ttu-id="d64ee-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-138">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-139">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-140">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-142">Office on Windows</span></span><br><span data-ttu-id="d64ee-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-144">- TaskPane</span></span><br><span data-ttu-id="d64ee-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-145">
        - Content</span></span><br><span data-ttu-id="d64ee-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d64ee-146">
        - Custom Functions</span></span><br><span data-ttu-id="d64ee-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="d64ee-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d64ee-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d64ee-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d64ee-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d64ee-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d64ee-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d64ee-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d64ee-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d64ee-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d64ee-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d64ee-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d64ee-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d64ee-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d64ee-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-161">
        - BindingEvents</span></span><br><span data-ttu-id="d64ee-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-162">
        - CompressedFile</span></span><br><span data-ttu-id="d64ee-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-163">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-164">
        - File</span></span><br><span data-ttu-id="d64ee-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-165">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-167">
        - Selection</span></span><br><span data-ttu-id="d64ee-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-168">
        - Settings</span></span><br><span data-ttu-id="d64ee-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-169">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-170">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-171">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-173">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-173">Office 2019 on Windows</span></span><br><span data-ttu-id="d64ee-174">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d64ee-175">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-175">- TaskPane</span></span><br><span data-ttu-id="d64ee-176">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-176">
        - Content</span></span><br><span data-ttu-id="d64ee-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d64ee-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d64ee-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d64ee-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d64ee-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d64ee-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d64ee-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d64ee-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d64ee-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d64ee-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d64ee-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-188">- BindingEvents</span></span><br><span data-ttu-id="d64ee-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-189">
        - CompressedFile</span></span><br><span data-ttu-id="d64ee-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-190">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-191">
        - File</span></span><br><span data-ttu-id="d64ee-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-192">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-194">
        - Selection</span></span><br><span data-ttu-id="d64ee-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-195">
        - Settings</span></span><br><span data-ttu-id="d64ee-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-196">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-197">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-198">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-200">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-200">Office 2016 on Windows</span></span><br><span data-ttu-id="d64ee-201">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d64ee-202">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-202">- TaskPane</span></span><br><span data-ttu-id="d64ee-203">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-203">
        - Content</span></span></td>
    <td><span data-ttu-id="d64ee-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d64ee-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d64ee-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d64ee-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d64ee-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-207">- BindingEvents</span></span><br><span data-ttu-id="d64ee-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-208">
        - CompressedFile</span></span><br><span data-ttu-id="d64ee-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-209">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-210">
        - File</span></span><br><span data-ttu-id="d64ee-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-211">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-213">
        - Selection</span></span><br><span data-ttu-id="d64ee-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-214">
        - Settings</span></span><br><span data-ttu-id="d64ee-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-215">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-216">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-217">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-219">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-219">Office 2013 on Windows</span></span><br><span data-ttu-id="d64ee-220">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d64ee-221">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-221">
        - TaskPane</span></span><br><span data-ttu-id="d64ee-222">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d64ee-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d64ee-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d64ee-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d64ee-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-225">
        - BindingEvents</span></span><br><span data-ttu-id="d64ee-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-226">
        - CompressedFile</span></span><br><span data-ttu-id="d64ee-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-227">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-228">
        - File</span></span><br><span data-ttu-id="d64ee-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-229">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-231">
        - Selection</span></span><br><span data-ttu-id="d64ee-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-232">
        - Settings</span></span><br><span data-ttu-id="d64ee-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-233">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-234">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-235">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-237">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d64ee-237">Office on iPad</span></span><br><span data-ttu-id="d64ee-238">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d64ee-239">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-239">- TaskPane</span></span><br><span data-ttu-id="d64ee-240">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-240">
        - Content</span></span></td>
    <td><span data-ttu-id="d64ee-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d64ee-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d64ee-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d64ee-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d64ee-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d64ee-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d64ee-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d64ee-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d64ee-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d64ee-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d64ee-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d64ee-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-253">- BindingEvents</span></span><br><span data-ttu-id="d64ee-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-254">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-255">
        - File</span></span><br><span data-ttu-id="d64ee-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-256">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-258">
        - Selection</span></span><br><span data-ttu-id="d64ee-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-259">
        - Settings</span></span><br><span data-ttu-id="d64ee-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-260">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-261">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-262">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-264">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-264">Office on Mac</span></span><br><span data-ttu-id="d64ee-265">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d64ee-266">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-266">- TaskPane</span></span><br><span data-ttu-id="d64ee-267">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-267">
        - Content</span></span><br><span data-ttu-id="d64ee-268">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d64ee-268">
        - Custom Functions</span></span><br><span data-ttu-id="d64ee-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d64ee-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d64ee-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d64ee-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d64ee-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d64ee-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d64ee-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d64ee-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d64ee-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d64ee-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d64ee-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d64ee-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d64ee-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d64ee-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-283">- BindingEvents</span></span><br><span data-ttu-id="d64ee-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-284">
        - CompressedFile</span></span><br><span data-ttu-id="d64ee-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-285">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-286">
        - File</span></span><br><span data-ttu-id="d64ee-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-287">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-289">
        - PdfFile</span></span><br><span data-ttu-id="d64ee-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-290">
        - Selection</span></span><br><span data-ttu-id="d64ee-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-291">
        - Settings</span></span><br><span data-ttu-id="d64ee-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-292">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-293">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-294">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-296">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-296">Office 2019 on Mac</span></span><br><span data-ttu-id="d64ee-297">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d64ee-298">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-298">- TaskPane</span></span><br><span data-ttu-id="d64ee-299">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-299">
        - Content</span></span><br><span data-ttu-id="d64ee-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d64ee-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d64ee-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d64ee-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d64ee-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d64ee-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d64ee-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d64ee-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d64ee-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d64ee-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d64ee-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-311">- BindingEvents</span></span><br><span data-ttu-id="d64ee-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-312">
        - CompressedFile</span></span><br><span data-ttu-id="d64ee-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-313">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-314">
        - File</span></span><br><span data-ttu-id="d64ee-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-315">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-317">
        - PdfFile</span></span><br><span data-ttu-id="d64ee-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-318">
        - Selection</span></span><br><span data-ttu-id="d64ee-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-319">
        - Settings</span></span><br><span data-ttu-id="d64ee-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-320">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-321">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-322">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-324">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-324">Office 2016 on Mac</span></span><br><span data-ttu-id="d64ee-325">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d64ee-326">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-326">- TaskPane</span></span><br><span data-ttu-id="d64ee-327">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-327">
        - Content</span></span></td>
    <td><span data-ttu-id="d64ee-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d64ee-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d64ee-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d64ee-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d64ee-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-331">- BindingEvents</span></span><br><span data-ttu-id="d64ee-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-332">
        - CompressedFile</span></span><br><span data-ttu-id="d64ee-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-333">
        - DocumentEvents</span></span><br><span data-ttu-id="d64ee-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-334">
        - File</span></span><br><span data-ttu-id="d64ee-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-335">
        - MatrixBindings</span></span><br><span data-ttu-id="d64ee-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-337">
        - PdfFile</span></span><br><span data-ttu-id="d64ee-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-338">
        - Selection</span></span><br><span data-ttu-id="d64ee-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-339">
        - Settings</span></span><br><span data-ttu-id="d64ee-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-340">
        - TableBindings</span></span><br><span data-ttu-id="d64ee-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-341">
        - TableCoercion</span></span><br><span data-ttu-id="d64ee-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-342">
        - TextBindings</span></span><br><span data-ttu-id="d64ee-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d64ee-344">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d64ee-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="d64ee-345">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d64ee-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d64ee-346">Платформа</span><span class="sxs-lookup"><span data-stu-id="d64ee-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d64ee-347">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d64ee-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d64ee-348">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d64ee-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d64ee-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d64ee-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-350">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d64ee-350">Office on the web</span></span></td>
    <td><span data-ttu-id="d64ee-351">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d64ee-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d64ee-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-353">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-353">Office on Windows</span></span><br><span data-ttu-id="d64ee-354">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d64ee-355">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d64ee-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d64ee-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-357">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-357">Office for Mac</span></span><br><span data-ttu-id="d64ee-358">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="d64ee-359">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d64ee-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d64ee-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="d64ee-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="d64ee-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d64ee-362">Платформа</span><span class="sxs-lookup"><span data-stu-id="d64ee-362">Platform</span></span></th>
    <th><span data-ttu-id="d64ee-363">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d64ee-363">Extension points</span></span></th>
    <th><span data-ttu-id="d64ee-364">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d64ee-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="d64ee-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d64ee-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-366">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d64ee-366">Office on the web</span></span><br><span data-ttu-id="d64ee-367">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="d64ee-367">(modern)</span></span></td>
    <td> <span data-ttu-id="d64ee-368">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-368">- Mail Read</span></span><br><span data-ttu-id="d64ee-369">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-369">
      - Mail Compose</span></span><br><span data-ttu-id="d64ee-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d64ee-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d64ee-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d64ee-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d64ee-379">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-380">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d64ee-380">Office on the web</span></span><br><span data-ttu-id="d64ee-381">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="d64ee-381">(classic)</span></span></td>
    <td> <span data-ttu-id="d64ee-382">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-382">- Mail Read</span></span><br><span data-ttu-id="d64ee-383">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-383">
      - Mail Compose</span></span><br><span data-ttu-id="d64ee-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d64ee-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d64ee-391">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-392">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-392">Office on Windows</span></span><br><span data-ttu-id="d64ee-393">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-394">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-394">- Mail Read</span></span><br><span data-ttu-id="d64ee-395">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-395">
      - Mail Compose</span></span><br><span data-ttu-id="d64ee-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d64ee-397">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="d64ee-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d64ee-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d64ee-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d64ee-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d64ee-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d64ee-406">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-407">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-407">Office 2019 on Windows</span></span><br><span data-ttu-id="d64ee-408">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-409">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-409">- Mail Read</span></span><br><span data-ttu-id="d64ee-410">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-410">
      - Mail Compose</span></span><br><span data-ttu-id="d64ee-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d64ee-412">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="d64ee-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d64ee-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d64ee-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d64ee-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d64ee-420">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-421">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-421">Office 2016 on Windows</span></span><br><span data-ttu-id="d64ee-422">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-423">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-423">- Mail Read</span></span><br><span data-ttu-id="d64ee-424">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-424">
      - Mail Compose</span></span><br><span data-ttu-id="d64ee-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d64ee-426">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="d64ee-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d64ee-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d64ee-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d64ee-431">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-432">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-432">Office 2013 on Windows</span></span><br><span data-ttu-id="d64ee-433">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-434">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-434">- Mail Read</span></span><br><span data-ttu-id="d64ee-435">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="d64ee-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="d64ee-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="d64ee-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d64ee-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d64ee-440">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-441">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="d64ee-441">Office on iOS</span></span><br><span data-ttu-id="d64ee-442">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-443">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-443">- Mail Read</span></span><br><span data-ttu-id="d64ee-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d64ee-450">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-451">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-451">Office on Mac</span></span><br><span data-ttu-id="d64ee-452">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-453">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-453">- Mail Read</span></span><br><span data-ttu-id="d64ee-454">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-454">
      - Mail Compose</span></span><br><span data-ttu-id="d64ee-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d64ee-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d64ee-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d64ee-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d64ee-464">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-465">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-465">Office 2019 on Mac</span></span><br><span data-ttu-id="d64ee-466">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-467">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-467">- Mail Read</span></span><br><span data-ttu-id="d64ee-468">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-468">
      - Mail Compose</span></span><br><span data-ttu-id="d64ee-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d64ee-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d64ee-476">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-477">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-477">Office 2016 on Mac</span></span><br><span data-ttu-id="d64ee-478">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-479">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-479">- Mail Read</span></span><br><span data-ttu-id="d64ee-480">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-480">
      - Mail Compose</span></span><br><span data-ttu-id="d64ee-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d64ee-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d64ee-488">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-489">Office для Android</span><span class="sxs-lookup"><span data-stu-id="d64ee-489">Office on Android</span></span><br><span data-ttu-id="d64ee-490">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-491">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d64ee-491">- Mail Read</span></span><br><span data-ttu-id="d64ee-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d64ee-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d64ee-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d64ee-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d64ee-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d64ee-498">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d64ee-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="d64ee-499">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d64ee-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d64ee-500">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="d64ee-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="d64ee-501">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="d64ee-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="d64ee-502">Word</span><span class="sxs-lookup"><span data-stu-id="d64ee-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d64ee-503">Платформа</span><span class="sxs-lookup"><span data-stu-id="d64ee-503">Platform</span></span></th>
    <th><span data-ttu-id="d64ee-504">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d64ee-504">Extension points</span></span></th>
    <th><span data-ttu-id="d64ee-505">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d64ee-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="d64ee-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d64ee-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-507">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d64ee-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="d64ee-508">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-508">- TaskPane</span></span><br><span data-ttu-id="d64ee-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d64ee-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d64ee-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d64ee-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d64ee-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d64ee-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-516">- BindingEvents</span></span><br><span data-ttu-id="d64ee-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-518">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-519">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-519">
         - File</span></span><br><span data-ttu-id="d64ee-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-521">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-524">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-525">
         - Selection</span></span><br><span data-ttu-id="d64ee-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-526">
         - Settings</span></span><br><span data-ttu-id="d64ee-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-527">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-528">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-529">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-530">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-532">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-532">Office on Windows</span></span><br><span data-ttu-id="d64ee-533">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-534">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-534">- TaskPane</span></span><br><span data-ttu-id="d64ee-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d64ee-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d64ee-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d64ee-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d64ee-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d64ee-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-542">- BindingEvents</span></span><br><span data-ttu-id="d64ee-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-543">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-545">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-546">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-546">
         - File</span></span><br><span data-ttu-id="d64ee-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-548">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-551">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-552">
         - Selection</span></span><br><span data-ttu-id="d64ee-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-553">
         - Settings</span></span><br><span data-ttu-id="d64ee-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-554">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-555">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-556">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-557">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-559">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-559">Office 2019 on Windows</span></span><br><span data-ttu-id="d64ee-560">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-561">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-561">- TaskPane</span></span><br><span data-ttu-id="d64ee-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d64ee-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d64ee-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d64ee-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-568">- BindingEvents</span></span><br><span data-ttu-id="d64ee-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-569">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-571">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-572">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-572">
         - File</span></span><br><span data-ttu-id="d64ee-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-574">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-577">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-578">
         - Selection</span></span><br><span data-ttu-id="d64ee-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-579">
         - Settings</span></span><br><span data-ttu-id="d64ee-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-580">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-581">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-582">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-583">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-585">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-585">Office 2016 on Windows</span></span><br><span data-ttu-id="d64ee-586">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-587">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d64ee-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d64ee-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d64ee-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-591">- BindingEvents</span></span><br><span data-ttu-id="d64ee-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-592">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-594">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-595">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-595">
         - File</span></span><br><span data-ttu-id="d64ee-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-597">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-600">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-601">
         - Selection</span></span><br><span data-ttu-id="d64ee-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-602">
         - Settings</span></span><br><span data-ttu-id="d64ee-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-603">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-604">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-605">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-606">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-608">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-608">Office 2013 on Windows</span></span><br><span data-ttu-id="d64ee-609">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-610">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d64ee-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d64ee-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-613">- BindingEvents</span></span><br><span data-ttu-id="d64ee-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-614">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-616">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-617">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-617">
         - File</span></span><br><span data-ttu-id="d64ee-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-619">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-622">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-623">
         - Selection</span></span><br><span data-ttu-id="d64ee-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-624">
         - Settings</span></span><br><span data-ttu-id="d64ee-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-625">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-626">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-627">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-628">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-630">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d64ee-630">Office on iPad</span></span><br><span data-ttu-id="d64ee-631">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-632">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d64ee-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d64ee-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d64ee-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d64ee-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-638">- BindingEvents</span></span><br><span data-ttu-id="d64ee-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-639">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-641">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-642">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-642">
         - File</span></span><br><span data-ttu-id="d64ee-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-644">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-647">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-648">
         - Selection</span></span><br><span data-ttu-id="d64ee-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-649">
         - Settings</span></span><br><span data-ttu-id="d64ee-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-650">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-651">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-652">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-653">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-655">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-655">Office on Mac</span></span><br><span data-ttu-id="d64ee-656">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-657">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-657">- TaskPane</span></span><br><span data-ttu-id="d64ee-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d64ee-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d64ee-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d64ee-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d64ee-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="d64ee-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-665">- BindingEvents</span></span><br><span data-ttu-id="d64ee-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-666">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-668">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-669">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-669">
         - File</span></span><br><span data-ttu-id="d64ee-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-671">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-674">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-675">
         - Selection</span></span><br><span data-ttu-id="d64ee-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-676">
         - Settings</span></span><br><span data-ttu-id="d64ee-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-677">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-678">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-679">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-680">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-682">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-682">Office 2019 on Mac</span></span><br><span data-ttu-id="d64ee-683">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-684">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-684">- TaskPane</span></span><br><span data-ttu-id="d64ee-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d64ee-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d64ee-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d64ee-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d64ee-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-691">- BindingEvents</span></span><br><span data-ttu-id="d64ee-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-692">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-694">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-695">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-695">
         - File</span></span><br><span data-ttu-id="d64ee-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-697">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-700">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-701">
         - Selection</span></span><br><span data-ttu-id="d64ee-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-702">
         - Settings</span></span><br><span data-ttu-id="d64ee-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-703">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-704">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-705">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-706">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-708">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-708">Office 2016 on Mac</span></span><br><span data-ttu-id="d64ee-709">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-710">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d64ee-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d64ee-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d64ee-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-714">- BindingEvents</span></span><br><span data-ttu-id="d64ee-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-715">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d64ee-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="d64ee-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-717">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-718">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d64ee-718">
         - File</span></span><br><span data-ttu-id="d64ee-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-720">
         - MatrixBindings</span></span><br><span data-ttu-id="d64ee-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="d64ee-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d64ee-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-723">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-724">
         - Selection</span></span><br><span data-ttu-id="d64ee-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d64ee-725">
         - Settings</span></span><br><span data-ttu-id="d64ee-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-726">
         - TableBindings</span></span><br><span data-ttu-id="d64ee-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-727">
         - TableCoercion</span></span><br><span data-ttu-id="d64ee-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d64ee-728">
         - TextBindings</span></span><br><span data-ttu-id="d64ee-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-729">
         - TextCoercion</span></span><br><span data-ttu-id="d64ee-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="d64ee-731">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d64ee-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d64ee-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d64ee-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d64ee-733">Платформа</span><span class="sxs-lookup"><span data-stu-id="d64ee-733">Platform</span></span></th>
    <th><span data-ttu-id="d64ee-734">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d64ee-734">Extension points</span></span></th>
    <th><span data-ttu-id="d64ee-735">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d64ee-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="d64ee-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d64ee-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-737">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d64ee-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="d64ee-738">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-738">- Content</span></span><br><span data-ttu-id="d64ee-739">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-739">
         - TaskPane</span></span><br><span data-ttu-id="d64ee-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d64ee-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d64ee-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d64ee-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-745">- ActiveView</span></span><br><span data-ttu-id="d64ee-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-746">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-747">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-748">
         - File</span></span><br><span data-ttu-id="d64ee-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-749">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-750">
         - Selection</span></span><br><span data-ttu-id="d64ee-751">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-751">
         - Settings</span></span><br><span data-ttu-id="d64ee-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-753">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-753">Office on Windows</span></span><br><span data-ttu-id="d64ee-754">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-755">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-755">- Content</span></span><br><span data-ttu-id="d64ee-756">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-756">
         - TaskPane</span></span><br><span data-ttu-id="d64ee-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d64ee-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d64ee-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d64ee-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-762">- ActiveView</span></span><br><span data-ttu-id="d64ee-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-763">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-764">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-765">
         - File</span></span><br><span data-ttu-id="d64ee-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-766">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-767">
         - Selection</span></span><br><span data-ttu-id="d64ee-768">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-768">
         - Settings</span></span><br><span data-ttu-id="d64ee-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-770">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-770">Office 2019 on Windows</span></span><br><span data-ttu-id="d64ee-771">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-772">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-772">- Content</span></span><br><span data-ttu-id="d64ee-773">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-773">
         - TaskPane</span></span><br><span data-ttu-id="d64ee-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-777">- ActiveView</span></span><br><span data-ttu-id="d64ee-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-778">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-779">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-780">
         - File</span></span><br><span data-ttu-id="d64ee-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-781">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-782">
         - Selection</span></span><br><span data-ttu-id="d64ee-783">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-783">
         - Settings</span></span><br><span data-ttu-id="d64ee-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-785">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-785">Office 2016 on Windows</span></span><br><span data-ttu-id="d64ee-786">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-787">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-787">- Content</span></span><br><span data-ttu-id="d64ee-788">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d64ee-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d64ee-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-791">- ActiveView</span></span><br><span data-ttu-id="d64ee-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-792">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-793">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-794">
         - File</span></span><br><span data-ttu-id="d64ee-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-795">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-796">
         - Selection</span></span><br><span data-ttu-id="d64ee-797">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-797">
         - Settings</span></span><br><span data-ttu-id="d64ee-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-799">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-799">Office 2013 on Windows</span></span><br><span data-ttu-id="d64ee-800">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-801">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-801">- Content</span></span><br><span data-ttu-id="d64ee-802">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="d64ee-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d64ee-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d64ee-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-805">- ActiveView</span></span><br><span data-ttu-id="d64ee-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-806">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-807">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-808">
         - File</span></span><br><span data-ttu-id="d64ee-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-809">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-810">
         - Selection</span></span><br><span data-ttu-id="d64ee-811">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-811">
         - Settings</span></span><br><span data-ttu-id="d64ee-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-813">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d64ee-813">Office on iPad</span></span><br><span data-ttu-id="d64ee-814">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-815">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-815">- Content</span></span><br><span data-ttu-id="d64ee-816">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d64ee-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-820">- ActiveView</span></span><br><span data-ttu-id="d64ee-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-821">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-822">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-823">
         - File</span></span><br><span data-ttu-id="d64ee-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-824">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-825">
         - Selection</span></span><br><span data-ttu-id="d64ee-826">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-826">
         - Settings</span></span><br><span data-ttu-id="d64ee-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-828">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-828">Office on Mac</span></span><br><span data-ttu-id="d64ee-829">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d64ee-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d64ee-830">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-830">- Content</span></span><br><span data-ttu-id="d64ee-831">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-831">
         - TaskPane</span></span><br><span data-ttu-id="d64ee-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d64ee-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d64ee-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d64ee-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-837">- ActiveView</span></span><br><span data-ttu-id="d64ee-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-838">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-839">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-840">
         - File</span></span><br><span data-ttu-id="d64ee-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-841">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-842">
         - Selection</span></span><br><span data-ttu-id="d64ee-843">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-843">
         - Settings</span></span><br><span data-ttu-id="d64ee-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-845">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-845">Office 2019 on Mac</span></span><br><span data-ttu-id="d64ee-846">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-847">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-847">- Content</span></span><br><span data-ttu-id="d64ee-848">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-848">
         - TaskPane</span></span><br><span data-ttu-id="d64ee-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-852">- ActiveView</span></span><br><span data-ttu-id="d64ee-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-853">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-854">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-855">
         - File</span></span><br><span data-ttu-id="d64ee-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-856">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-857">
         - Selection</span></span><br><span data-ttu-id="d64ee-858">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-858">
         - Settings</span></span><br><span data-ttu-id="d64ee-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-860">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-860">Office 2016 on Mac</span></span><br><span data-ttu-id="d64ee-861">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-862">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-862">- Content</span></span><br><span data-ttu-id="d64ee-863">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d64ee-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d64ee-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d64ee-866">- ActiveView</span></span><br><span data-ttu-id="d64ee-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-867">
         - CompressedFile</span></span><br><span data-ttu-id="d64ee-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-868">
         - DocumentEvents</span></span><br><span data-ttu-id="d64ee-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="d64ee-869">
         - File</span></span><br><span data-ttu-id="d64ee-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d64ee-870">
         - PdfFile</span></span><br><span data-ttu-id="d64ee-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-871">
         - Selection</span></span><br><span data-ttu-id="d64ee-872">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-872">
         - Settings</span></span><br><span data-ttu-id="d64ee-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d64ee-874">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d64ee-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="d64ee-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="d64ee-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d64ee-876">Платформа</span><span class="sxs-lookup"><span data-stu-id="d64ee-876">Platform</span></span></th>
    <th><span data-ttu-id="d64ee-877">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d64ee-877">Extension points</span></span></th>
    <th><span data-ttu-id="d64ee-878">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d64ee-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="d64ee-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d64ee-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-880">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d64ee-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="d64ee-881">- Контент</span><span class="sxs-lookup"><span data-stu-id="d64ee-881">- Content</span></span><br><span data-ttu-id="d64ee-882">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-882">
         - TaskPane</span></span><br><span data-ttu-id="d64ee-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d64ee-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d64ee-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d64ee-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d64ee-887">- DocumentEvents</span></span><br><span data-ttu-id="d64ee-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="d64ee-889">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d64ee-889">
         - Settings</span></span><br><span data-ttu-id="d64ee-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="d64ee-891">Project</span><span class="sxs-lookup"><span data-stu-id="d64ee-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d64ee-892">Платформа</span><span class="sxs-lookup"><span data-stu-id="d64ee-892">Platform</span></span></th>
    <th><span data-ttu-id="d64ee-893">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d64ee-893">Extension points</span></span></th>
    <th><span data-ttu-id="d64ee-894">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d64ee-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="d64ee-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d64ee-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-896">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-896">Office 2019 on Windows</span></span><br><span data-ttu-id="d64ee-897">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-898">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-900">- Selection</span></span><br><span data-ttu-id="d64ee-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-902">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-902">Office 2016 on Windows</span></span><br><span data-ttu-id="d64ee-903">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-904">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-906">- Selection</span></span><br><span data-ttu-id="d64ee-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d64ee-908">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d64ee-908">Office 2013 on Windows</span></span><br><span data-ttu-id="d64ee-909">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d64ee-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d64ee-910">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d64ee-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d64ee-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d64ee-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d64ee-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="d64ee-912">- Selection</span></span><br><span data-ttu-id="d64ee-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d64ee-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="d64ee-914">См. также</span><span class="sxs-lookup"><span data-stu-id="d64ee-914">See also</span></span>

- [<span data-ttu-id="d64ee-915">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d64ee-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d64ee-916">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="d64ee-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d64ee-917">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="d64ee-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="d64ee-918">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="d64ee-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="d64ee-919">Справочная документация по API</span><span class="sxs-lookup"><span data-stu-id="d64ee-919">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="d64ee-920">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="d64ee-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="d64ee-921">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="d64ee-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="d64ee-922">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="d64ee-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="d64ee-923">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d64ee-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="d64ee-924">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d64ee-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="d64ee-925">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d64ee-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="d64ee-926">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d64ee-926">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)