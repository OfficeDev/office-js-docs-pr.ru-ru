---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: a3c580f32ad7cd384309a9b53e55ea488a470a90
ms.sourcegitcommit: f781d7cfd980cd866d6d1d00c5b9d16c8a4b7f9b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/20/2019
ms.locfileid: "37053328"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="16731-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="16731-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="16731-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="16731-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="16731-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="16731-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="16731-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="16731-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="16731-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="16731-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="16731-108">Excel</span><span class="sxs-lookup"><span data-stu-id="16731-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="16731-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="16731-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="16731-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="16731-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="16731-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="16731-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="16731-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="16731-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="16731-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="16731-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-114">- TaskPane</span></span><br><span data-ttu-id="16731-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-115">
        - Content</span></span><br><span data-ttu-id="16731-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="16731-116">
        - Custom Functions</span></span><br><span data-ttu-id="16731-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="16731-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="16731-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16731-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16731-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16731-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16731-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16731-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16731-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16731-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16731-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16731-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="16731-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="16731-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="16731-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-128">
        - BindingEvents</span></span><br><span data-ttu-id="16731-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-129">
        - CompressedFile</span></span><br><span data-ttu-id="16731-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-130">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-131">
        - File</span></span><br><span data-ttu-id="16731-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-132">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-134">
        - Selection</span></span><br><span data-ttu-id="16731-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-135">
        - Settings</span></span><br><span data-ttu-id="16731-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-136">
        - TableBindings</span></span><br><span data-ttu-id="16731-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-137">
        - TableCoercion</span></span><br><span data-ttu-id="16731-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-138">
        - TextBindings</span></span><br><span data-ttu-id="16731-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-140">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-140">Office on Windows</span></span><br><span data-ttu-id="16731-141">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-142">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-142">- TaskPane</span></span><br><span data-ttu-id="16731-143">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-143">
        - Content</span></span><br><span data-ttu-id="16731-144">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="16731-144">
        - Custom Functions</span></span><br><span data-ttu-id="16731-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="16731-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="16731-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16731-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16731-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16731-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16731-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16731-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16731-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16731-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16731-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16731-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="16731-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="16731-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16731-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="16731-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-158">
        - BindingEvents</span></span><br><span data-ttu-id="16731-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-159">
        - CompressedFile</span></span><br><span data-ttu-id="16731-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-160">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-161">
        - File</span></span><br><span data-ttu-id="16731-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-162">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-164">
        - Selection</span></span><br><span data-ttu-id="16731-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-165">
        - Settings</span></span><br><span data-ttu-id="16731-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-166">
        - TableBindings</span></span><br><span data-ttu-id="16731-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-167">
        - TableCoercion</span></span><br><span data-ttu-id="16731-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-168">
        - TextBindings</span></span><br><span data-ttu-id="16731-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-170">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-170">Office 2019 on Windows</span></span><br><span data-ttu-id="16731-171">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16731-172">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-172">- TaskPane</span></span><br><span data-ttu-id="16731-173">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-173">
        - Content</span></span><br><span data-ttu-id="16731-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="16731-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16731-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16731-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16731-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16731-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16731-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16731-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16731-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16731-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16731-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16731-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-185">- BindingEvents</span></span><br><span data-ttu-id="16731-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-186">
        - CompressedFile</span></span><br><span data-ttu-id="16731-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-187">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-188">
        - File</span></span><br><span data-ttu-id="16731-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-189">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-191">
        - Selection</span></span><br><span data-ttu-id="16731-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-192">
        - Settings</span></span><br><span data-ttu-id="16731-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-193">
        - TableBindings</span></span><br><span data-ttu-id="16731-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-194">
        - TableCoercion</span></span><br><span data-ttu-id="16731-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-195">
        - TextBindings</span></span><br><span data-ttu-id="16731-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-197">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-197">Office 2016 on Windows</span></span><br><span data-ttu-id="16731-198">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16731-199">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-199">- TaskPane</span></span><br><span data-ttu-id="16731-200">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-200">
        - Content</span></span></td>
    <td><span data-ttu-id="16731-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16731-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="16731-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="16731-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16731-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-204">- BindingEvents</span></span><br><span data-ttu-id="16731-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-205">
        - CompressedFile</span></span><br><span data-ttu-id="16731-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-206">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-207">
        - File</span></span><br><span data-ttu-id="16731-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-208">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-210">
        - Selection</span></span><br><span data-ttu-id="16731-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-211">
        - Settings</span></span><br><span data-ttu-id="16731-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-212">
        - TableBindings</span></span><br><span data-ttu-id="16731-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-213">
        - TableCoercion</span></span><br><span data-ttu-id="16731-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-214">
        - TextBindings</span></span><br><span data-ttu-id="16731-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-216">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-216">Office 2013 on Windows</span></span><br><span data-ttu-id="16731-217">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16731-218">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-218">
        - TaskPane</span></span><br><span data-ttu-id="16731-219">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="16731-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16731-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16731-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16731-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-222">
        - BindingEvents</span></span><br><span data-ttu-id="16731-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-223">
        - CompressedFile</span></span><br><span data-ttu-id="16731-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-224">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-225">
        - File</span></span><br><span data-ttu-id="16731-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-226">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-228">
        - Selection</span></span><br><span data-ttu-id="16731-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-229">
        - Settings</span></span><br><span data-ttu-id="16731-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-230">
        - TableBindings</span></span><br><span data-ttu-id="16731-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-231">
        - TableCoercion</span></span><br><span data-ttu-id="16731-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-232">
        - TextBindings</span></span><br><span data-ttu-id="16731-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-234">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="16731-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="16731-235">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="16731-236">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-236">- TaskPane</span></span><br><span data-ttu-id="16731-237">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-237">
        - Content</span></span></td>
    <td><span data-ttu-id="16731-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16731-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16731-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16731-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16731-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16731-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16731-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16731-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16731-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16731-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="16731-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="16731-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16731-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-249">- BindingEvents</span></span><br><span data-ttu-id="16731-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-250">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-251">
        - File</span></span><br><span data-ttu-id="16731-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-252">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-254">
        - Selection</span></span><br><span data-ttu-id="16731-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-255">
        - Settings</span></span><br><span data-ttu-id="16731-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-256">
        - TableBindings</span></span><br><span data-ttu-id="16731-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-257">
        - TableCoercion</span></span><br><span data-ttu-id="16731-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-258">
        - TextBindings</span></span><br><span data-ttu-id="16731-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-260">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-260">Office apps on Mac</span></span><br><span data-ttu-id="16731-261">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="16731-262">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-262">- TaskPane</span></span><br><span data-ttu-id="16731-263">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-263">
        - Content</span></span><br><span data-ttu-id="16731-264">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="16731-264">
        - Custom Functions</span></span><br><span data-ttu-id="16731-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="16731-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16731-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16731-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16731-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16731-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16731-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16731-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16731-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16731-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16731-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="16731-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="16731-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16731-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="16731-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-278">- BindingEvents</span></span><br><span data-ttu-id="16731-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-279">
        - CompressedFile</span></span><br><span data-ttu-id="16731-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-280">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-281">
        - File</span></span><br><span data-ttu-id="16731-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-282">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-284">
        - PdfFile</span></span><br><span data-ttu-id="16731-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-285">
        - Selection</span></span><br><span data-ttu-id="16731-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-286">
        - Settings</span></span><br><span data-ttu-id="16731-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-287">
        - TableBindings</span></span><br><span data-ttu-id="16731-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-288">
        - TableCoercion</span></span><br><span data-ttu-id="16731-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-289">
        - TextBindings</span></span><br><span data-ttu-id="16731-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-291">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-291">Office 2019 for Mac</span></span><br><span data-ttu-id="16731-292">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16731-293">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-293">- TaskPane</span></span><br><span data-ttu-id="16731-294">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-294">
        - Content</span></span><br><span data-ttu-id="16731-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="16731-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16731-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16731-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16731-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16731-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16731-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16731-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16731-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16731-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16731-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16731-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-306">- BindingEvents</span></span><br><span data-ttu-id="16731-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-307">
        - CompressedFile</span></span><br><span data-ttu-id="16731-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-308">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-309">
        - File</span></span><br><span data-ttu-id="16731-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-310">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-312">
        - PdfFile</span></span><br><span data-ttu-id="16731-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-313">
        - Selection</span></span><br><span data-ttu-id="16731-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-314">
        - Settings</span></span><br><span data-ttu-id="16731-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-315">
        - TableBindings</span></span><br><span data-ttu-id="16731-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-316">
        - TableCoercion</span></span><br><span data-ttu-id="16731-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-317">
        - TextBindings</span></span><br><span data-ttu-id="16731-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-319">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-319">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="16731-320">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16731-321">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-321">- TaskPane</span></span><br><span data-ttu-id="16731-322">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="16731-322">
        - Content</span></span></td>
    <td><span data-ttu-id="16731-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16731-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="16731-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="16731-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16731-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-326">- BindingEvents</span></span><br><span data-ttu-id="16731-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-327">
        - CompressedFile</span></span><br><span data-ttu-id="16731-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-328">
        - DocumentEvents</span></span><br><span data-ttu-id="16731-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="16731-329">
        - File</span></span><br><span data-ttu-id="16731-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-330">
        - MatrixBindings</span></span><br><span data-ttu-id="16731-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="16731-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-332">
        - PdfFile</span></span><br><span data-ttu-id="16731-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-333">
        - Selection</span></span><br><span data-ttu-id="16731-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-334">
        - Settings</span></span><br><span data-ttu-id="16731-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-335">
        - TableBindings</span></span><br><span data-ttu-id="16731-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-336">
        - TableCoercion</span></span><br><span data-ttu-id="16731-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-337">
        - TextBindings</span></span><br><span data-ttu-id="16731-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="16731-339">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="16731-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="16731-340">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="16731-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="16731-341">Платформа</span><span class="sxs-lookup"><span data-stu-id="16731-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="16731-342">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="16731-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="16731-343">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="16731-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="16731-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="16731-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-345">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="16731-345">Office on the web</span></span></td>
    <td><span data-ttu-id="16731-346">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="16731-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="16731-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-348">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-348">Office on Windows</span></span><br><span data-ttu-id="16731-349">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="16731-350">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="16731-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="16731-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-352">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-352">Office for Mac</span></span><br><span data-ttu-id="16731-353">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="16731-354">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="16731-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="16731-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="16731-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="16731-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16731-357">Платформа</span><span class="sxs-lookup"><span data-stu-id="16731-357">Platform</span></span></th>
    <th><span data-ttu-id="16731-358">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="16731-358">Extension points</span></span></th>
    <th><span data-ttu-id="16731-359">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="16731-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="16731-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="16731-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-361">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="16731-361">Office on the web</span></span><br><span data-ttu-id="16731-362">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="16731-362">Modern</span></span></td>
    <td> <span data-ttu-id="16731-363">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-363">- Mail Read</span></span><br><span data-ttu-id="16731-364">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-364">
      - Mail Compose</span></span><br><span data-ttu-id="16731-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16731-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="16731-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="16731-373">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-374">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="16731-374">Office on the web</span></span><br><span data-ttu-id="16731-375">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="16731-375">(classic)</span></span></td>
    <td> <span data-ttu-id="16731-376">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-376">- Mail Read</span></span><br><span data-ttu-id="16731-377">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-377">
      - Mail Compose</span></span><br><span data-ttu-id="16731-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16731-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="16731-385">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-386">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-386">Office on Windows</span></span><br><span data-ttu-id="16731-387">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-387">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-388">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-388">- Mail Read</span></span><br><span data-ttu-id="16731-389">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-389">
      - Mail Compose</span></span><br><span data-ttu-id="16731-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="16731-391">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="16731-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="16731-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16731-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="16731-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="16731-399">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-400">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-400">Office 2019 on Windows</span></span><br><span data-ttu-id="16731-401">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-402">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-402">- Mail Read</span></span><br><span data-ttu-id="16731-403">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-403">
      - Mail Compose</span></span><br><span data-ttu-id="16731-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="16731-405">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="16731-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="16731-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16731-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="16731-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="16731-413">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-414">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-414">Office 2016 on Windows</span></span><br><span data-ttu-id="16731-415">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-416">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-416">- Mail Read</span></span><br><span data-ttu-id="16731-417">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-417">
      - Mail Compose</span></span><br><span data-ttu-id="16731-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="16731-419">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="16731-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="16731-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="16731-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="16731-424">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-425">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-425">Office 2013 on Windows</span></span><br><span data-ttu-id="16731-426">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-427">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-427">- Mail Read</span></span><br><span data-ttu-id="16731-428">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="16731-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="16731-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="16731-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="16731-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="16731-433">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-434">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="16731-434">Office apps on iOS</span></span><br><span data-ttu-id="16731-435">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-435">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-436">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-436">- Mail Read</span></span><br><span data-ttu-id="16731-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="16731-443">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-444">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-444">Office apps on Mac</span></span><br><span data-ttu-id="16731-445">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-445">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-446">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-446">- Mail Read</span></span><br><span data-ttu-id="16731-447">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-447">
      - Mail Compose</span></span><br><span data-ttu-id="16731-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16731-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="16731-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16731-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="16731-456">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-457">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-457">Office 2019 for Mac</span></span><br><span data-ttu-id="16731-458">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-459">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-459">- Mail Read</span></span><br><span data-ttu-id="16731-460">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-460">
      - Mail Compose</span></span><br><span data-ttu-id="16731-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16731-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="16731-468">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-469">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-469">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="16731-470">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-471">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-471">- Mail Read</span></span><br><span data-ttu-id="16731-472">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="16731-472">
      - Mail Compose</span></span><br><span data-ttu-id="16731-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16731-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16731-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="16731-480">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-481">Office для Android</span><span class="sxs-lookup"><span data-stu-id="16731-481">Office apps on Android</span></span><br><span data-ttu-id="16731-482">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-482">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-483">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="16731-483">- Mail Read</span></span><br><span data-ttu-id="16731-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16731-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16731-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16731-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16731-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16731-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16731-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="16731-490">Недоступно</span><span class="sxs-lookup"><span data-stu-id="16731-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="16731-491">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="16731-491">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="16731-492">Word</span><span class="sxs-lookup"><span data-stu-id="16731-492">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16731-493">Платформа</span><span class="sxs-lookup"><span data-stu-id="16731-493">Platform</span></span></th>
    <th><span data-ttu-id="16731-494">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="16731-494">Extension points</span></span></th>
    <th><span data-ttu-id="16731-495">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="16731-495">API requirement sets</span></span></th>
    <th><span data-ttu-id="16731-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="16731-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-497">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="16731-497">Office on the web</span></span></td>
    <td> <span data-ttu-id="16731-498">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-498">- TaskPane</span></span><br><span data-ttu-id="16731-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16731-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16731-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16731-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16731-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16731-506">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-506">- BindingEvents</span></span><br><span data-ttu-id="16731-507">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-507">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-508">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-508">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-509">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-509">
         - File</span></span><br><span data-ttu-id="16731-510">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-510">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-511">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-511">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-512">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-512">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-513">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-513">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-514">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-514">
         - PdfFile</span></span><br><span data-ttu-id="16731-515">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-515">
         - Selection</span></span><br><span data-ttu-id="16731-516">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-516">
         - Settings</span></span><br><span data-ttu-id="16731-517">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-517">
         - TableBindings</span></span><br><span data-ttu-id="16731-518">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-518">
         - TableCoercion</span></span><br><span data-ttu-id="16731-519">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-519">
         - TextBindings</span></span><br><span data-ttu-id="16731-520">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-520">
         - TextCoercion</span></span><br><span data-ttu-id="16731-521">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-521">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-522">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-522">Office on Windows</span></span><br><span data-ttu-id="16731-523">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-523">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-524">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-524">- TaskPane</span></span><br><span data-ttu-id="16731-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16731-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16731-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16731-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16731-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16731-532">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-532">- BindingEvents</span></span><br><span data-ttu-id="16731-533">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-533">
         - CompressedFile</span></span><br><span data-ttu-id="16731-534">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-534">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-535">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-535">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-536">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-536">
         - File</span></span><br><span data-ttu-id="16731-537">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-537">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-538">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-538">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-539">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-539">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-540">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-540">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-541">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-541">
         - PdfFile</span></span><br><span data-ttu-id="16731-542">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-542">
         - Selection</span></span><br><span data-ttu-id="16731-543">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-543">
         - Settings</span></span><br><span data-ttu-id="16731-544">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-544">
         - TableBindings</span></span><br><span data-ttu-id="16731-545">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-545">
         - TableCoercion</span></span><br><span data-ttu-id="16731-546">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-546">
         - TextBindings</span></span><br><span data-ttu-id="16731-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-547">
         - TextCoercion</span></span><br><span data-ttu-id="16731-548">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-548">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-549">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-549">Office 2019 on Windows</span></span><br><span data-ttu-id="16731-550">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-550">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-551">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-551">- TaskPane</span></span><br><span data-ttu-id="16731-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16731-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16731-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16731-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-558">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-558">- BindingEvents</span></span><br><span data-ttu-id="16731-559">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-559">
         - CompressedFile</span></span><br><span data-ttu-id="16731-560">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-560">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-561">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-561">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-562">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-562">
         - File</span></span><br><span data-ttu-id="16731-563">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-563">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-564">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-564">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-565">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-565">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-566">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-566">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-567">
         - PdfFile</span></span><br><span data-ttu-id="16731-568">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-568">
         - Selection</span></span><br><span data-ttu-id="16731-569">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-569">
         - Settings</span></span><br><span data-ttu-id="16731-570">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-570">
         - TableBindings</span></span><br><span data-ttu-id="16731-571">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-571">
         - TableCoercion</span></span><br><span data-ttu-id="16731-572">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-572">
         - TextBindings</span></span><br><span data-ttu-id="16731-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-573">
         - TextCoercion</span></span><br><span data-ttu-id="16731-574">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-574">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-575">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-575">Office 2016 on Windows</span></span><br><span data-ttu-id="16731-576">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-576">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-577">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-577">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16731-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="16731-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="16731-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-581">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-581">- BindingEvents</span></span><br><span data-ttu-id="16731-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-582">
         - CompressedFile</span></span><br><span data-ttu-id="16731-583">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-583">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-584">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-584">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-585">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-585">
         - File</span></span><br><span data-ttu-id="16731-586">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-586">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-587">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-587">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-588">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-588">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-589">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-589">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-590">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-590">
         - PdfFile</span></span><br><span data-ttu-id="16731-591">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-591">
         - Selection</span></span><br><span data-ttu-id="16731-592">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-592">
         - Settings</span></span><br><span data-ttu-id="16731-593">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-593">
         - TableBindings</span></span><br><span data-ttu-id="16731-594">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-594">
         - TableCoercion</span></span><br><span data-ttu-id="16731-595">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-595">
         - TextBindings</span></span><br><span data-ttu-id="16731-596">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-596">
         - TextCoercion</span></span><br><span data-ttu-id="16731-597">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-597">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-598">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-598">Office 2013 on Windows</span></span><br><span data-ttu-id="16731-599">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-599">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-600">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-600">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16731-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16731-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-603">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-603">- BindingEvents</span></span><br><span data-ttu-id="16731-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-604">
         - CompressedFile</span></span><br><span data-ttu-id="16731-605">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-605">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-606">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-607">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-607">
         - File</span></span><br><span data-ttu-id="16731-608">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-608">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-609">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-609">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-610">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-610">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-611">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-611">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-612">
         - PdfFile</span></span><br><span data-ttu-id="16731-613">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-613">
         - Selection</span></span><br><span data-ttu-id="16731-614">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-614">
         - Settings</span></span><br><span data-ttu-id="16731-615">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-615">
         - TableBindings</span></span><br><span data-ttu-id="16731-616">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-616">
         - TableCoercion</span></span><br><span data-ttu-id="16731-617">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-617">
         - TextBindings</span></span><br><span data-ttu-id="16731-618">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-618">
         - TextCoercion</span></span><br><span data-ttu-id="16731-619">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-619">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-620">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="16731-620">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="16731-621">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-621">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-622">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-622">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16731-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16731-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16731-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="16731-628">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-628">- BindingEvents</span></span><br><span data-ttu-id="16731-629">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-629">
         - CompressedFile</span></span><br><span data-ttu-id="16731-630">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-630">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-631">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-631">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-632">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-632">
         - File</span></span><br><span data-ttu-id="16731-633">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-633">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-634">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-634">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-635">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-635">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-636">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-636">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-637">
         - PdfFile</span></span><br><span data-ttu-id="16731-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-638">
         - Selection</span></span><br><span data-ttu-id="16731-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-639">
         - Settings</span></span><br><span data-ttu-id="16731-640">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-640">
         - TableBindings</span></span><br><span data-ttu-id="16731-641">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-641">
         - TableCoercion</span></span><br><span data-ttu-id="16731-642">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-642">
         - TextBindings</span></span><br><span data-ttu-id="16731-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-643">
         - TextCoercion</span></span><br><span data-ttu-id="16731-644">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-644">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-645">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-645">Office apps on Mac</span></span><br><span data-ttu-id="16731-646">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-646">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-647">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-647">- TaskPane</span></span><br><span data-ttu-id="16731-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16731-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16731-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16731-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16731-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="16731-655">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-655">- BindingEvents</span></span><br><span data-ttu-id="16731-656">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-656">
         - CompressedFile</span></span><br><span data-ttu-id="16731-657">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-657">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-658">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-658">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-659">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-659">
         - File</span></span><br><span data-ttu-id="16731-660">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-660">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-661">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-661">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-662">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-662">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-663">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-663">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-664">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-664">
         - PdfFile</span></span><br><span data-ttu-id="16731-665">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-665">
         - Selection</span></span><br><span data-ttu-id="16731-666">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-666">
         - Settings</span></span><br><span data-ttu-id="16731-667">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-667">
         - TableBindings</span></span><br><span data-ttu-id="16731-668">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-668">
         - TableCoercion</span></span><br><span data-ttu-id="16731-669">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-669">
         - TextBindings</span></span><br><span data-ttu-id="16731-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-670">
         - TextCoercion</span></span><br><span data-ttu-id="16731-671">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-671">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-672">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-672">Office 2019 for Mac</span></span><br><span data-ttu-id="16731-673">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-673">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-674">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-674">- TaskPane</span></span><br><span data-ttu-id="16731-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16731-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16731-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16731-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16731-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="16731-681">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-681">- BindingEvents</span></span><br><span data-ttu-id="16731-682">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-682">
         - CompressedFile</span></span><br><span data-ttu-id="16731-683">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-683">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-684">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-684">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-685">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-685">
         - File</span></span><br><span data-ttu-id="16731-686">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-686">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-687">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-687">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-688">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-688">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-689">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-689">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-690">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-690">
         - PdfFile</span></span><br><span data-ttu-id="16731-691">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-691">
         - Selection</span></span><br><span data-ttu-id="16731-692">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-692">
         - Settings</span></span><br><span data-ttu-id="16731-693">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-693">
         - TableBindings</span></span><br><span data-ttu-id="16731-694">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-694">
         - TableCoercion</span></span><br><span data-ttu-id="16731-695">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-695">
         - TextBindings</span></span><br><span data-ttu-id="16731-696">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-696">
         - TextCoercion</span></span><br><span data-ttu-id="16731-697">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-697">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-698">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-698">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="16731-699">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-699">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-700">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-700">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16731-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="16731-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="16731-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-704">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16731-704">- BindingEvents</span></span><br><span data-ttu-id="16731-705">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-705">
         - CompressedFile</span></span><br><span data-ttu-id="16731-706">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16731-706">
         - CustomXmlParts</span></span><br><span data-ttu-id="16731-707">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-707">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-708">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="16731-708">
         - File</span></span><br><span data-ttu-id="16731-709">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-709">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-710">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16731-710">
         - MatrixBindings</span></span><br><span data-ttu-id="16731-711">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-711">
         - MatrixCoercion</span></span><br><span data-ttu-id="16731-712">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-712">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16731-713">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-713">
         - PdfFile</span></span><br><span data-ttu-id="16731-714">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-714">
         - Selection</span></span><br><span data-ttu-id="16731-715">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16731-715">
         - Settings</span></span><br><span data-ttu-id="16731-716">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16731-716">
         - TableBindings</span></span><br><span data-ttu-id="16731-717">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-717">
         - TableCoercion</span></span><br><span data-ttu-id="16731-718">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16731-718">
         - TextBindings</span></span><br><span data-ttu-id="16731-719">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-719">
         - TextCoercion</span></span><br><span data-ttu-id="16731-720">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16731-720">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="16731-721">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="16731-721">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="16731-722">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="16731-722">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16731-723">Платформа</span><span class="sxs-lookup"><span data-stu-id="16731-723">Platform</span></span></th>
    <th><span data-ttu-id="16731-724">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="16731-724">Extension points</span></span></th>
    <th><span data-ttu-id="16731-725">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="16731-725">API requirement sets</span></span></th>
    <th><span data-ttu-id="16731-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="16731-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-727">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="16731-727">Office on the web</span></span></td>
    <td> <span data-ttu-id="16731-728">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-728">- Content</span></span><br><span data-ttu-id="16731-729">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-729">
         - TaskPane</span></span><br><span data-ttu-id="16731-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="16731-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16731-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16731-735">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-735">- ActiveView</span></span><br><span data-ttu-id="16731-736">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-736">
         - CompressedFile</span></span><br><span data-ttu-id="16731-737">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-737">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-738">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-738">
         - File</span></span><br><span data-ttu-id="16731-739">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-739">
         - PdfFile</span></span><br><span data-ttu-id="16731-740">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-740">
         - Selection</span></span><br><span data-ttu-id="16731-741">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-741">
         - Settings</span></span><br><span data-ttu-id="16731-742">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-742">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-743">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-743">Office on Windows</span></span><br><span data-ttu-id="16731-744">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-744">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-745">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-745">- Content</span></span><br><span data-ttu-id="16731-746">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-746">
         - TaskPane</span></span><br><span data-ttu-id="16731-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="16731-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16731-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16731-752">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-752">- ActiveView</span></span><br><span data-ttu-id="16731-753">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-753">
         - CompressedFile</span></span><br><span data-ttu-id="16731-754">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-754">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-755">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-755">
         - File</span></span><br><span data-ttu-id="16731-756">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-756">
         - PdfFile</span></span><br><span data-ttu-id="16731-757">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-757">
         - Selection</span></span><br><span data-ttu-id="16731-758">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-758">
         - Settings</span></span><br><span data-ttu-id="16731-759">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-759">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-760">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-760">Office 2019 on Windows</span></span><br><span data-ttu-id="16731-761">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-761">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-762">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-762">- Content</span></span><br><span data-ttu-id="16731-763">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-763">
         - TaskPane</span></span><br><span data-ttu-id="16731-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-767">- ActiveView</span></span><br><span data-ttu-id="16731-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-768">
         - CompressedFile</span></span><br><span data-ttu-id="16731-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-769">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-770">
         - File</span></span><br><span data-ttu-id="16731-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-771">
         - PdfFile</span></span><br><span data-ttu-id="16731-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-772">
         - Selection</span></span><br><span data-ttu-id="16731-773">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-773">
         - Settings</span></span><br><span data-ttu-id="16731-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-775">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-775">Office 2016 on Windows</span></span><br><span data-ttu-id="16731-776">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-776">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-777">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-777">- Content</span></span><br><span data-ttu-id="16731-778">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-778">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16731-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16731-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-781">- ActiveView</span></span><br><span data-ttu-id="16731-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-782">
         - CompressedFile</span></span><br><span data-ttu-id="16731-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-783">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-784">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-784">
         - File</span></span><br><span data-ttu-id="16731-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-785">
         - PdfFile</span></span><br><span data-ttu-id="16731-786">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-786">
         - Selection</span></span><br><span data-ttu-id="16731-787">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-787">
         - Settings</span></span><br><span data-ttu-id="16731-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-789">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-789">Office 2013 on Windows</span></span><br><span data-ttu-id="16731-790">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-791">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-791">- Content</span></span><br><span data-ttu-id="16731-792">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-792">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="16731-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16731-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16731-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-795">- ActiveView</span></span><br><span data-ttu-id="16731-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-796">
         - CompressedFile</span></span><br><span data-ttu-id="16731-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-797">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-798">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-798">
         - File</span></span><br><span data-ttu-id="16731-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-799">
         - PdfFile</span></span><br><span data-ttu-id="16731-800">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-800">
         - Selection</span></span><br><span data-ttu-id="16731-801">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-801">
         - Settings</span></span><br><span data-ttu-id="16731-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-803">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="16731-803">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="16731-804">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-804">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-805">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-805">- Content</span></span><br><span data-ttu-id="16731-806">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="16731-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-810">- ActiveView</span></span><br><span data-ttu-id="16731-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-811">
         - CompressedFile</span></span><br><span data-ttu-id="16731-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-812">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-813">
         - File</span></span><br><span data-ttu-id="16731-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-814">
         - PdfFile</span></span><br><span data-ttu-id="16731-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-815">
         - Selection</span></span><br><span data-ttu-id="16731-816">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-816">
         - Settings</span></span><br><span data-ttu-id="16731-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-818">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-818">Office apps on Mac</span></span><br><span data-ttu-id="16731-819">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="16731-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16731-820">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-820">- Content</span></span><br><span data-ttu-id="16731-821">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-821">
         - TaskPane</span></span><br><span data-ttu-id="16731-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="16731-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16731-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16731-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16731-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-827">- ActiveView</span></span><br><span data-ttu-id="16731-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-828">
         - CompressedFile</span></span><br><span data-ttu-id="16731-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-829">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-830">
         - File</span></span><br><span data-ttu-id="16731-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-831">
         - PdfFile</span></span><br><span data-ttu-id="16731-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-832">
         - Selection</span></span><br><span data-ttu-id="16731-833">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-833">
         - Settings</span></span><br><span data-ttu-id="16731-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-835">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-835">Office 2019 for Mac</span></span><br><span data-ttu-id="16731-836">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-836">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-837">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-837">- Content</span></span><br><span data-ttu-id="16731-838">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-838">
         - TaskPane</span></span><br><span data-ttu-id="16731-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-842">- ActiveView</span></span><br><span data-ttu-id="16731-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-843">
         - CompressedFile</span></span><br><span data-ttu-id="16731-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-844">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-845">
         - File</span></span><br><span data-ttu-id="16731-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-846">
         - PdfFile</span></span><br><span data-ttu-id="16731-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-847">
         - Selection</span></span><br><span data-ttu-id="16731-848">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-848">
         - Settings</span></span><br><span data-ttu-id="16731-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-850">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-850">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="16731-851">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-851">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-852">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-852">- Content</span></span><br><span data-ttu-id="16731-853">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-853">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16731-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16731-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16731-856">- ActiveView</span></span><br><span data-ttu-id="16731-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16731-857">
         - CompressedFile</span></span><br><span data-ttu-id="16731-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-858">
         - DocumentEvents</span></span><br><span data-ttu-id="16731-859">
         - File</span><span class="sxs-lookup"><span data-stu-id="16731-859">
         - File</span></span><br><span data-ttu-id="16731-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16731-860">
         - PdfFile</span></span><br><span data-ttu-id="16731-861">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16731-861">
         - Selection</span></span><br><span data-ttu-id="16731-862">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-862">
         - Settings</span></span><br><span data-ttu-id="16731-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-863">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="16731-864">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="16731-864">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="16731-865">OneNote</span><span class="sxs-lookup"><span data-stu-id="16731-865">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16731-866">Платформа</span><span class="sxs-lookup"><span data-stu-id="16731-866">Platform</span></span></th>
    <th><span data-ttu-id="16731-867">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="16731-867">Extension points</span></span></th>
    <th><span data-ttu-id="16731-868">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="16731-868">API requirement sets</span></span></th>
    <th><span data-ttu-id="16731-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="16731-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-870">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="16731-870">Office on the web</span></span></td>
    <td> <span data-ttu-id="16731-871">- Контент</span><span class="sxs-lookup"><span data-stu-id="16731-871">- Content</span></span><br><span data-ttu-id="16731-872">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-872">
         - TaskPane</span></span><br><span data-ttu-id="16731-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="16731-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16731-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="16731-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16731-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-877">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16731-877">- DocumentEvents</span></span><br><span data-ttu-id="16731-878">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-878">
         - HtmlCoercion</span></span><br><span data-ttu-id="16731-879">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="16731-879">
         - Settings</span></span><br><span data-ttu-id="16731-880">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-880">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="16731-881">Project</span><span class="sxs-lookup"><span data-stu-id="16731-881">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16731-882">Платформа</span><span class="sxs-lookup"><span data-stu-id="16731-882">Platform</span></span></th>
    <th><span data-ttu-id="16731-883">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="16731-883">Extension points</span></span></th>
    <th><span data-ttu-id="16731-884">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="16731-884">API requirement sets</span></span></th>
    <th><span data-ttu-id="16731-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="16731-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-886">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-886">Office 2019 on Windows</span></span><br><span data-ttu-id="16731-887">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-888">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-890">- Selection</span><span class="sxs-lookup"><span data-stu-id="16731-890">- Selection</span></span><br><span data-ttu-id="16731-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-891">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-892">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-892">Office 2016 on Windows</span></span><br><span data-ttu-id="16731-893">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-893">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-894">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-894">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-896">- Selection</span><span class="sxs-lookup"><span data-stu-id="16731-896">- Selection</span></span><br><span data-ttu-id="16731-897">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-897">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16731-898">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="16731-898">Office 2013 on Windows</span></span><br><span data-ttu-id="16731-899">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="16731-899">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16731-900">- Область задач</span><span class="sxs-lookup"><span data-stu-id="16731-900">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16731-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16731-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="16731-902">- Selection</span><span class="sxs-lookup"><span data-stu-id="16731-902">- Selection</span></span><br><span data-ttu-id="16731-903">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16731-903">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="16731-904">См. также</span><span class="sxs-lookup"><span data-stu-id="16731-904">See also</span></span>

- [<span data-ttu-id="16731-905">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="16731-905">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="16731-906">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="16731-906">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="16731-907">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="16731-907">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="16731-908">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="16731-908">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="16731-909">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="16731-909">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="16731-910">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="16731-910">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="16731-911">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="16731-911">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="16731-912">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="16731-912">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="16731-913">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="16731-913">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="16731-914">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="16731-914">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="16731-915">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="16731-915">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
