---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: d88f7c1b9daa201d9b6bc5cfa69ac3125bf127b1
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/11/2019
ms.locfileid: "35630538"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d2774-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d2774-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d2774-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="d2774-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="d2774-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="d2774-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="d2774-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="d2774-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="d2774-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="d2774-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="d2774-108">Excel</span><span class="sxs-lookup"><span data-stu-id="d2774-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d2774-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="d2774-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d2774-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d2774-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d2774-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d2774-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d2774-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d2774-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d2774-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="d2774-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-114">- TaskPane</span></span><br><span data-ttu-id="d2774-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-115">
        - Content</span></span><br><span data-ttu-id="d2774-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d2774-116">
        - Custom Functions</span></span><br><span data-ttu-id="d2774-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="d2774-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d2774-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d2774-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d2774-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d2774-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d2774-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d2774-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d2774-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d2774-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d2774-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d2774-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d2774-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d2774-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d2774-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-130">
        - BindingEvents</span></span><br><span data-ttu-id="d2774-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-131">
        - CompressedFile</span></span><br><span data-ttu-id="d2774-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-132">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-133">
        - File</span></span><br><span data-ttu-id="d2774-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-134">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-136">
        - Selection</span></span><br><span data-ttu-id="d2774-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-137">
        - Settings</span></span><br><span data-ttu-id="d2774-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-138">
        - TableBindings</span></span><br><span data-ttu-id="d2774-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-139">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-140">
        - TextBindings</span></span><br><span data-ttu-id="d2774-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-142">Office on Windows</span></span><br><span data-ttu-id="d2774-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-144">- TaskPane</span></span><br><span data-ttu-id="d2774-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-145">
        - Content</span></span><br><span data-ttu-id="d2774-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d2774-146">
        - Custom Functions</span></span><br><span data-ttu-id="d2774-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="d2774-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d2774-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d2774-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d2774-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d2774-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d2774-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d2774-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d2774-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d2774-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d2774-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d2774-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d2774-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d2774-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d2774-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-160">
        - BindingEvents</span></span><br><span data-ttu-id="d2774-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-161">
        - CompressedFile</span></span><br><span data-ttu-id="d2774-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-162">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-163">
        - File</span></span><br><span data-ttu-id="d2774-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-164">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-166">
        - Selection</span></span><br><span data-ttu-id="d2774-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-167">
        - Settings</span></span><br><span data-ttu-id="d2774-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-168">
        - TableBindings</span></span><br><span data-ttu-id="d2774-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-169">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-170">
        - TextBindings</span></span><br><span data-ttu-id="d2774-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-172">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-172">Office 2019 on Windows</span></span><br><span data-ttu-id="d2774-173">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d2774-174">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-174">- TaskPane</span></span><br><span data-ttu-id="d2774-175">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-175">
        - Content</span></span><br><span data-ttu-id="d2774-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d2774-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d2774-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d2774-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d2774-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d2774-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d2774-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d2774-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d2774-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d2774-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d2774-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d2774-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-187">- BindingEvents</span></span><br><span data-ttu-id="d2774-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-188">
        - CompressedFile</span></span><br><span data-ttu-id="d2774-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-189">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-190">
        - File</span></span><br><span data-ttu-id="d2774-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-191">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-193">
        - Selection</span></span><br><span data-ttu-id="d2774-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-194">
        - Settings</span></span><br><span data-ttu-id="d2774-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-195">
        - TableBindings</span></span><br><span data-ttu-id="d2774-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-196">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-197">
        - TextBindings</span></span><br><span data-ttu-id="d2774-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-199">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-199">Office 2016 on Windows</span></span><br><span data-ttu-id="d2774-200">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d2774-201">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-201">- TaskPane</span></span><br><span data-ttu-id="d2774-202">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-202">
        - Content</span></span></td>
    <td><span data-ttu-id="d2774-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d2774-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d2774-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d2774-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d2774-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-206">- BindingEvents</span></span><br><span data-ttu-id="d2774-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-207">
        - CompressedFile</span></span><br><span data-ttu-id="d2774-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-208">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-209">
        - File</span></span><br><span data-ttu-id="d2774-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-210">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-212">
        - Selection</span></span><br><span data-ttu-id="d2774-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-213">
        - Settings</span></span><br><span data-ttu-id="d2774-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-214">
        - TableBindings</span></span><br><span data-ttu-id="d2774-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-215">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-216">
        - TextBindings</span></span><br><span data-ttu-id="d2774-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-218">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-218">Office 2013 on Windows</span></span><br><span data-ttu-id="d2774-219">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d2774-220">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-220">
        - TaskPane</span></span><br><span data-ttu-id="d2774-221">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d2774-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d2774-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d2774-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d2774-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-224">
        - BindingEvents</span></span><br><span data-ttu-id="d2774-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-225">
        - CompressedFile</span></span><br><span data-ttu-id="d2774-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-226">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-227">
        - File</span></span><br><span data-ttu-id="d2774-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-228">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-230">
        - Selection</span></span><br><span data-ttu-id="d2774-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-231">
        - Settings</span></span><br><span data-ttu-id="d2774-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-232">
        - TableBindings</span></span><br><span data-ttu-id="d2774-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-233">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-234">
        - TextBindings</span></span><br><span data-ttu-id="d2774-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-236">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d2774-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="d2774-237">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d2774-238">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-238">- TaskPane</span></span><br><span data-ttu-id="d2774-239">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-239">
        - Content</span></span><br><span data-ttu-id="d2774-240">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d2774-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d2774-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d2774-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d2774-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d2774-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d2774-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d2774-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d2774-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d2774-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d2774-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d2774-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d2774-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d2774-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d2774-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-252">- BindingEvents</span></span><br><span data-ttu-id="d2774-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-253">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-254">
        - File</span></span><br><span data-ttu-id="d2774-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-255">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-257">
        - Selection</span></span><br><span data-ttu-id="d2774-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-258">
        - Settings</span></span><br><span data-ttu-id="d2774-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-259">
        - TableBindings</span></span><br><span data-ttu-id="d2774-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-260">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-261">
        - TextBindings</span></span><br><span data-ttu-id="d2774-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-263">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-263">Office apps on Mac</span></span><br><span data-ttu-id="d2774-264">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d2774-265">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-265">- TaskPane</span></span><br><span data-ttu-id="d2774-266">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-266">
        - Content</span></span><br><span data-ttu-id="d2774-267">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d2774-267">
        - Custom Functions</span></span><br><span data-ttu-id="d2774-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d2774-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d2774-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d2774-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d2774-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d2774-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d2774-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d2774-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d2774-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d2774-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d2774-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d2774-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d2774-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d2774-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-281">- BindingEvents</span></span><br><span data-ttu-id="d2774-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-282">
        - CompressedFile</span></span><br><span data-ttu-id="d2774-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-283">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-284">
        - File</span></span><br><span data-ttu-id="d2774-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-285">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-287">
        - PdfFile</span></span><br><span data-ttu-id="d2774-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-288">
        - Selection</span></span><br><span data-ttu-id="d2774-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-289">
        - Settings</span></span><br><span data-ttu-id="d2774-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-290">
        - TableBindings</span></span><br><span data-ttu-id="d2774-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-291">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-292">
        - TextBindings</span></span><br><span data-ttu-id="d2774-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-294">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-294">Office 2019 for Mac</span></span><br><span data-ttu-id="d2774-295">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d2774-296">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-296">- TaskPane</span></span><br><span data-ttu-id="d2774-297">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-297">
        - Content</span></span><br><span data-ttu-id="d2774-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d2774-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d2774-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d2774-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d2774-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d2774-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d2774-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d2774-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d2774-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d2774-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d2774-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d2774-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-309">- BindingEvents</span></span><br><span data-ttu-id="d2774-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-310">
        - CompressedFile</span></span><br><span data-ttu-id="d2774-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-311">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-312">
        - File</span></span><br><span data-ttu-id="d2774-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-313">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-315">
        - PdfFile</span></span><br><span data-ttu-id="d2774-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-316">
        - Selection</span></span><br><span data-ttu-id="d2774-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-317">
        - Settings</span></span><br><span data-ttu-id="d2774-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-318">
        - TableBindings</span></span><br><span data-ttu-id="d2774-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-319">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-320">
        - TextBindings</span></span><br><span data-ttu-id="d2774-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-322">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-322">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="d2774-323">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d2774-324">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-324">- TaskPane</span></span><br><span data-ttu-id="d2774-325">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-325">
        - Content</span></span></td>
    <td><span data-ttu-id="d2774-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d2774-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d2774-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d2774-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d2774-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-329">- BindingEvents</span></span><br><span data-ttu-id="d2774-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-330">
        - CompressedFile</span></span><br><span data-ttu-id="d2774-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-331">
        - DocumentEvents</span></span><br><span data-ttu-id="d2774-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="d2774-332">
        - File</span></span><br><span data-ttu-id="d2774-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-333">
        - MatrixBindings</span></span><br><span data-ttu-id="d2774-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="d2774-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-335">
        - PdfFile</span></span><br><span data-ttu-id="d2774-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-336">
        - Selection</span></span><br><span data-ttu-id="d2774-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-337">
        - Settings</span></span><br><span data-ttu-id="d2774-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-338">
        - TableBindings</span></span><br><span data-ttu-id="d2774-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-339">
        - TableCoercion</span></span><br><span data-ttu-id="d2774-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-340">
        - TextBindings</span></span><br><span data-ttu-id="d2774-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d2774-342">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d2774-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="d2774-343">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d2774-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d2774-344">Платформа</span><span class="sxs-lookup"><span data-stu-id="d2774-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d2774-345">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d2774-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d2774-346">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d2774-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d2774-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d2774-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-348">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d2774-348">Office on the web</span></span></td>
    <td><span data-ttu-id="d2774-349">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d2774-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d2774-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-351">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-351">Office on Windows</span></span><br><span data-ttu-id="d2774-352">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d2774-353">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d2774-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d2774-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-355">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-355">Office for Mac</span></span><br><span data-ttu-id="d2774-356">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="d2774-357">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="d2774-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d2774-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="d2774-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="d2774-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d2774-360">Платформа</span><span class="sxs-lookup"><span data-stu-id="d2774-360">Platform</span></span></th>
    <th><span data-ttu-id="d2774-361">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d2774-361">Extension points</span></span></th>
    <th><span data-ttu-id="d2774-362">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d2774-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="d2774-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d2774-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-364">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d2774-364">Office on the web</span></span><br><span data-ttu-id="d2774-365">(новый)</span><span class="sxs-lookup"><span data-stu-id="d2774-365">New</span></span></td>
    <td> <span data-ttu-id="d2774-366">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-366">- Mail Read</span></span><br><span data-ttu-id="d2774-367">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-367">
      - Mail Compose</span></span><br><span data-ttu-id="d2774-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d2774-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d2774-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d2774-376">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-377">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d2774-377">Office on the web</span></span><br><span data-ttu-id="d2774-378">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="d2774-378">(classic)</span></span></td>
    <td> <span data-ttu-id="d2774-379">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-379">- Mail Read</span></span><br><span data-ttu-id="d2774-380">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-380">
      - Mail Compose</span></span><br><span data-ttu-id="d2774-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d2774-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d2774-388">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-389">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-389">Office on Windows</span></span><br><span data-ttu-id="d2774-390">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-391">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-391">- Mail Read</span></span><br><span data-ttu-id="d2774-392">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-392">
      - Mail Compose</span></span><br><span data-ttu-id="d2774-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d2774-394">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="d2774-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d2774-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d2774-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d2774-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d2774-402">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-403">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-403">Office 2019 on Windows</span></span><br><span data-ttu-id="d2774-404">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-405">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-405">- Mail Read</span></span><br><span data-ttu-id="d2774-406">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-406">
      - Mail Compose</span></span><br><span data-ttu-id="d2774-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d2774-408">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="d2774-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d2774-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d2774-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d2774-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d2774-416">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-417">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-417">Office 2016 on Windows</span></span><br><span data-ttu-id="d2774-418">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-419">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-419">- Mail Read</span></span><br><span data-ttu-id="d2774-420">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-420">
      - Mail Compose</span></span><br><span data-ttu-id="d2774-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d2774-422">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="d2774-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d2774-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d2774-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d2774-427">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-428">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-428">Office 2013 on Windows</span></span><br><span data-ttu-id="d2774-429">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-430">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-430">- Mail Read</span></span><br><span data-ttu-id="d2774-431">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="d2774-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="d2774-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="d2774-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d2774-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d2774-436">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-437">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="d2774-437">Office apps on iOS</span></span><br><span data-ttu-id="d2774-438">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-439">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-439">- Mail Read</span></span><br><span data-ttu-id="d2774-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d2774-446">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-447">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-447">Office apps on Mac</span></span><br><span data-ttu-id="d2774-448">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-449">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-449">- Mail Read</span></span><br><span data-ttu-id="d2774-450">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-450">
      - Mail Compose</span></span><br><span data-ttu-id="d2774-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d2774-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d2774-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d2774-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d2774-459">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-460">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-460">Office 2019 for Mac</span></span><br><span data-ttu-id="d2774-461">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-462">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-462">- Mail Read</span></span><br><span data-ttu-id="d2774-463">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-463">
      - Mail Compose</span></span><br><span data-ttu-id="d2774-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d2774-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d2774-471">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-472">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-472">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="d2774-473">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-474">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-474">- Mail Read</span></span><br><span data-ttu-id="d2774-475">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d2774-475">
      - Mail Compose</span></span><br><span data-ttu-id="d2774-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d2774-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d2774-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d2774-483">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-484">Office для Android</span><span class="sxs-lookup"><span data-stu-id="d2774-484">Office apps on Android</span></span><br><span data-ttu-id="d2774-485">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-486">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d2774-486">- Mail Read</span></span><br><span data-ttu-id="d2774-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d2774-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d2774-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d2774-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d2774-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d2774-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d2774-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d2774-493">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d2774-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="d2774-494">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d2774-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="d2774-495">Word</span><span class="sxs-lookup"><span data-stu-id="d2774-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d2774-496">Платформа</span><span class="sxs-lookup"><span data-stu-id="d2774-496">Platform</span></span></th>
    <th><span data-ttu-id="d2774-497">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d2774-497">Extension points</span></span></th>
    <th><span data-ttu-id="d2774-498">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d2774-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="d2774-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d2774-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-500">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d2774-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="d2774-501">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-501">- TaskPane</span></span><br><span data-ttu-id="d2774-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d2774-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d2774-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d2774-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d2774-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-509">- BindingEvents</span></span><br><span data-ttu-id="d2774-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-511">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-512">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-512">
         - File</span></span><br><span data-ttu-id="d2774-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-514">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-517">
         - PdfFile</span></span><br><span data-ttu-id="d2774-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-518">
         - Selection</span></span><br><span data-ttu-id="d2774-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-519">
         - Settings</span></span><br><span data-ttu-id="d2774-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-520">
         - TableBindings</span></span><br><span data-ttu-id="d2774-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-521">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-522">
         - TextBindings</span></span><br><span data-ttu-id="d2774-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-523">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-525">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-525">Office on Windows</span></span><br><span data-ttu-id="d2774-526">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-527">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-527">- TaskPane</span></span><br><span data-ttu-id="d2774-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d2774-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d2774-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d2774-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d2774-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-535">- BindingEvents</span></span><br><span data-ttu-id="d2774-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-536">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-538">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-539">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-539">
         - File</span></span><br><span data-ttu-id="d2774-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-541">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-544">
         - PdfFile</span></span><br><span data-ttu-id="d2774-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-545">
         - Selection</span></span><br><span data-ttu-id="d2774-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-546">
         - Settings</span></span><br><span data-ttu-id="d2774-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-547">
         - TableBindings</span></span><br><span data-ttu-id="d2774-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-548">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-549">
         - TextBindings</span></span><br><span data-ttu-id="d2774-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-550">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-552">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-552">Office 2019 on Windows</span></span><br><span data-ttu-id="d2774-553">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-554">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-554">- TaskPane</span></span><br><span data-ttu-id="d2774-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d2774-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d2774-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d2774-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-561">- BindingEvents</span></span><br><span data-ttu-id="d2774-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-562">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-564">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-565">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-565">
         - File</span></span><br><span data-ttu-id="d2774-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-567">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-570">
         - PdfFile</span></span><br><span data-ttu-id="d2774-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-571">
         - Selection</span></span><br><span data-ttu-id="d2774-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-572">
         - Settings</span></span><br><span data-ttu-id="d2774-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-573">
         - TableBindings</span></span><br><span data-ttu-id="d2774-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-574">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-575">
         - TextBindings</span></span><br><span data-ttu-id="d2774-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-576">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-578">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-578">Office 2016 on Windows</span></span><br><span data-ttu-id="d2774-579">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-580">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d2774-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d2774-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d2774-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-584">- BindingEvents</span></span><br><span data-ttu-id="d2774-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-585">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-587">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-588">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-588">
         - File</span></span><br><span data-ttu-id="d2774-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-590">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-593">
         - PdfFile</span></span><br><span data-ttu-id="d2774-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-594">
         - Selection</span></span><br><span data-ttu-id="d2774-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-595">
         - Settings</span></span><br><span data-ttu-id="d2774-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-596">
         - TableBindings</span></span><br><span data-ttu-id="d2774-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-597">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-598">
         - TextBindings</span></span><br><span data-ttu-id="d2774-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-599">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-601">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-601">Office 2013 on Windows</span></span><br><span data-ttu-id="d2774-602">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-603">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d2774-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d2774-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-606">- BindingEvents</span></span><br><span data-ttu-id="d2774-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-607">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-609">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-610">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-610">
         - File</span></span><br><span data-ttu-id="d2774-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-612">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-615">
         - PdfFile</span></span><br><span data-ttu-id="d2774-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-616">
         - Selection</span></span><br><span data-ttu-id="d2774-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-617">
         - Settings</span></span><br><span data-ttu-id="d2774-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-618">
         - TableBindings</span></span><br><span data-ttu-id="d2774-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-619">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-620">
         - TextBindings</span></span><br><span data-ttu-id="d2774-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-621">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-623">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d2774-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="d2774-624">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-625">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d2774-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d2774-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d2774-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d2774-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-631">- BindingEvents</span></span><br><span data-ttu-id="d2774-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-632">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-634">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-635">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-635">
         - File</span></span><br><span data-ttu-id="d2774-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-637">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-640">
         - PdfFile</span></span><br><span data-ttu-id="d2774-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-641">
         - Selection</span></span><br><span data-ttu-id="d2774-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-642">
         - Settings</span></span><br><span data-ttu-id="d2774-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-643">
         - TableBindings</span></span><br><span data-ttu-id="d2774-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-644">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-645">
         - TextBindings</span></span><br><span data-ttu-id="d2774-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-646">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-648">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-648">Office apps on Mac</span></span><br><span data-ttu-id="d2774-649">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-650">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-650">- TaskPane</span></span><br><span data-ttu-id="d2774-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d2774-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d2774-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d2774-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="d2774-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-658">- BindingEvents</span></span><br><span data-ttu-id="d2774-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-659">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-661">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-662">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-662">
         - File</span></span><br><span data-ttu-id="d2774-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-664">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-667">
         - PdfFile</span></span><br><span data-ttu-id="d2774-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-668">
         - Selection</span></span><br><span data-ttu-id="d2774-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-669">
         - Settings</span></span><br><span data-ttu-id="d2774-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-670">
         - TableBindings</span></span><br><span data-ttu-id="d2774-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-671">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-672">
         - TextBindings</span></span><br><span data-ttu-id="d2774-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-673">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-675">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-675">Office 2019 for Mac</span></span><br><span data-ttu-id="d2774-676">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-677">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-677">- TaskPane</span></span><br><span data-ttu-id="d2774-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d2774-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d2774-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d2774-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d2774-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d2774-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-684">- BindingEvents</span></span><br><span data-ttu-id="d2774-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-685">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-687">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-688">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-688">
         - File</span></span><br><span data-ttu-id="d2774-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-690">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-693">
         - PdfFile</span></span><br><span data-ttu-id="d2774-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-694">
         - Selection</span></span><br><span data-ttu-id="d2774-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-695">
         - Settings</span></span><br><span data-ttu-id="d2774-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-696">
         - TableBindings</span></span><br><span data-ttu-id="d2774-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-697">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-698">
         - TextBindings</span></span><br><span data-ttu-id="d2774-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-699">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-701">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-701">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="d2774-702">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-703">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d2774-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d2774-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d2774-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-707">- BindingEvents</span></span><br><span data-ttu-id="d2774-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-708">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d2774-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="d2774-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-710">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-711">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d2774-711">
         - File</span></span><br><span data-ttu-id="d2774-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-713">
         - MatrixBindings</span></span><br><span data-ttu-id="d2774-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="d2774-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d2774-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-716">
         - PdfFile</span></span><br><span data-ttu-id="d2774-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-717">
         - Selection</span></span><br><span data-ttu-id="d2774-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d2774-718">
         - Settings</span></span><br><span data-ttu-id="d2774-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-719">
         - TableBindings</span></span><br><span data-ttu-id="d2774-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-720">
         - TableCoercion</span></span><br><span data-ttu-id="d2774-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d2774-721">
         - TextBindings</span></span><br><span data-ttu-id="d2774-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-722">
         - TextCoercion</span></span><br><span data-ttu-id="d2774-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d2774-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="d2774-724">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d2774-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d2774-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d2774-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d2774-726">Платформа</span><span class="sxs-lookup"><span data-stu-id="d2774-726">Platform</span></span></th>
    <th><span data-ttu-id="d2774-727">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d2774-727">Extension points</span></span></th>
    <th><span data-ttu-id="d2774-728">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d2774-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="d2774-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d2774-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-730">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d2774-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="d2774-731">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-731">- Content</span></span><br><span data-ttu-id="d2774-732">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-732">
         - TaskPane</span></span><br><span data-ttu-id="d2774-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d2774-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-737">- ActiveView</span></span><br><span data-ttu-id="d2774-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-738">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-739">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-740">
         - File</span></span><br><span data-ttu-id="d2774-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-741">
         - PdfFile</span></span><br><span data-ttu-id="d2774-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-742">
         - Selection</span></span><br><span data-ttu-id="d2774-743">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-743">
         - Settings</span></span><br><span data-ttu-id="d2774-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-745">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-745">Office on Windows</span></span><br><span data-ttu-id="d2774-746">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-747">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-747">- Content</span></span><br><span data-ttu-id="d2774-748">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-748">
         - TaskPane</span></span><br><span data-ttu-id="d2774-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d2774-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-753">- ActiveView</span></span><br><span data-ttu-id="d2774-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-754">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-755">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-756">
         - File</span></span><br><span data-ttu-id="d2774-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-757">
         - PdfFile</span></span><br><span data-ttu-id="d2774-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-758">
         - Selection</span></span><br><span data-ttu-id="d2774-759">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-759">
         - Settings</span></span><br><span data-ttu-id="d2774-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-761">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-761">Office 2019 on Windows</span></span><br><span data-ttu-id="d2774-762">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-763">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-763">- Content</span></span><br><span data-ttu-id="d2774-764">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-764">
         - TaskPane</span></span><br><span data-ttu-id="d2774-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-768">- ActiveView</span></span><br><span data-ttu-id="d2774-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-769">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-770">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-771">
         - File</span></span><br><span data-ttu-id="d2774-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-772">
         - PdfFile</span></span><br><span data-ttu-id="d2774-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-773">
         - Selection</span></span><br><span data-ttu-id="d2774-774">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-774">
         - Settings</span></span><br><span data-ttu-id="d2774-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-776">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-776">Office 2016 on Windows</span></span><br><span data-ttu-id="d2774-777">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-778">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-778">- Content</span></span><br><span data-ttu-id="d2774-779">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d2774-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d2774-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-782">- ActiveView</span></span><br><span data-ttu-id="d2774-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-783">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-784">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-785">
         - File</span></span><br><span data-ttu-id="d2774-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-786">
         - PdfFile</span></span><br><span data-ttu-id="d2774-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-787">
         - Selection</span></span><br><span data-ttu-id="d2774-788">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-788">
         - Settings</span></span><br><span data-ttu-id="d2774-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-790">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-790">Office 2013 on Windows</span></span><br><span data-ttu-id="d2774-791">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-792">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-792">- Content</span></span><br><span data-ttu-id="d2774-793">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="d2774-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d2774-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d2774-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-796">- ActiveView</span></span><br><span data-ttu-id="d2774-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-797">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-798">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-799">
         - File</span></span><br><span data-ttu-id="d2774-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-800">
         - PdfFile</span></span><br><span data-ttu-id="d2774-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-801">
         - Selection</span></span><br><span data-ttu-id="d2774-802">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-802">
         - Settings</span></span><br><span data-ttu-id="d2774-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-804">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d2774-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="d2774-805">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-806">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-806">- Content</span></span><br><span data-ttu-id="d2774-807">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-810">- ActiveView</span></span><br><span data-ttu-id="d2774-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-811">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-812">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-813">
         - File</span></span><br><span data-ttu-id="d2774-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-814">
         - PdfFile</span></span><br><span data-ttu-id="d2774-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-815">
         - Selection</span></span><br><span data-ttu-id="d2774-816">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-816">
         - Settings</span></span><br><span data-ttu-id="d2774-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-818">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-818">Office apps on Mac</span></span><br><span data-ttu-id="d2774-819">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d2774-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d2774-820">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-820">- Content</span></span><br><span data-ttu-id="d2774-821">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-821">
         - TaskPane</span></span><br><span data-ttu-id="d2774-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d2774-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d2774-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d2774-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-826">- ActiveView</span></span><br><span data-ttu-id="d2774-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-827">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-828">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-829">
         - File</span></span><br><span data-ttu-id="d2774-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-830">
         - PdfFile</span></span><br><span data-ttu-id="d2774-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-831">
         - Selection</span></span><br><span data-ttu-id="d2774-832">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-832">
         - Settings</span></span><br><span data-ttu-id="d2774-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-834">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-834">Office 2019 for Mac</span></span><br><span data-ttu-id="d2774-835">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-836">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-836">- Content</span></span><br><span data-ttu-id="d2774-837">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-837">
         - TaskPane</span></span><br><span data-ttu-id="d2774-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-841">- ActiveView</span></span><br><span data-ttu-id="d2774-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-842">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-843">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-844">
         - File</span></span><br><span data-ttu-id="d2774-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-845">
         - PdfFile</span></span><br><span data-ttu-id="d2774-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-846">
         - Selection</span></span><br><span data-ttu-id="d2774-847">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-847">
         - Settings</span></span><br><span data-ttu-id="d2774-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-849">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-849">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="d2774-850">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-851">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-851">- Content</span></span><br><span data-ttu-id="d2774-852">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d2774-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d2774-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d2774-855">- ActiveView</span></span><br><span data-ttu-id="d2774-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d2774-856">
         - CompressedFile</span></span><br><span data-ttu-id="d2774-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-857">
         - DocumentEvents</span></span><br><span data-ttu-id="d2774-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="d2774-858">
         - File</span></span><br><span data-ttu-id="d2774-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d2774-859">
         - PdfFile</span></span><br><span data-ttu-id="d2774-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-860">
         - Selection</span></span><br><span data-ttu-id="d2774-861">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-861">
         - Settings</span></span><br><span data-ttu-id="d2774-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d2774-863">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d2774-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="d2774-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="d2774-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d2774-865">Платформа</span><span class="sxs-lookup"><span data-stu-id="d2774-865">Platform</span></span></th>
    <th><span data-ttu-id="d2774-866">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d2774-866">Extension points</span></span></th>
    <th><span data-ttu-id="d2774-867">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d2774-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="d2774-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d2774-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-869">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="d2774-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="d2774-870">- Контент</span><span class="sxs-lookup"><span data-stu-id="d2774-870">- Content</span></span><br><span data-ttu-id="d2774-871">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-871">
         - TaskPane</span></span><br><span data-ttu-id="d2774-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d2774-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d2774-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d2774-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d2774-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d2774-876">- DocumentEvents</span></span><br><span data-ttu-id="d2774-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="d2774-878">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d2774-878">
         - Settings</span></span><br><span data-ttu-id="d2774-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="d2774-880">Project</span><span class="sxs-lookup"><span data-stu-id="d2774-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d2774-881">Платформа</span><span class="sxs-lookup"><span data-stu-id="d2774-881">Platform</span></span></th>
    <th><span data-ttu-id="d2774-882">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d2774-882">Extension points</span></span></th>
    <th><span data-ttu-id="d2774-883">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d2774-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="d2774-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d2774-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-885">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-885">Office 2019 on Windows</span></span><br><span data-ttu-id="d2774-886">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-887">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-889">- Selection</span></span><br><span data-ttu-id="d2774-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-891">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-891">Office 2016 on Windows</span></span><br><span data-ttu-id="d2774-892">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-893">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-895">- Selection</span></span><br><span data-ttu-id="d2774-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d2774-897">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d2774-897">Office 2013 on Windows</span></span><br><span data-ttu-id="d2774-898">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="d2774-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d2774-899">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d2774-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d2774-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d2774-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d2774-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="d2774-901">- Selection</span></span><br><span data-ttu-id="d2774-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d2774-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="d2774-903">См. также</span><span class="sxs-lookup"><span data-stu-id="d2774-903">See also</span></span>

- [<span data-ttu-id="d2774-904">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d2774-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d2774-905">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="d2774-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="d2774-906">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="d2774-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="d2774-907">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="d2774-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="d2774-908">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="d2774-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="d2774-909">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="d2774-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="d2774-910">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="d2774-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="d2774-911">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="d2774-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="d2774-912">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d2774-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="d2774-913">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d2774-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="d2774-914">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="d2774-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
