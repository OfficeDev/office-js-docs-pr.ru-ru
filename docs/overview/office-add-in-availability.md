---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 10/09/2019
localization_priority: Priority
ms.openlocfilehash: 28d63866a03bcae99829d3a6b6c6198059a92bdc
ms.sourcegitcommit: 4d9f3e177b0bcd62804d5045f52b03e441af244f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2019
ms.locfileid: "37440152"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="beb6e-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="beb6e-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="beb6e-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="beb6e-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="beb6e-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="beb6e-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="beb6e-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="beb6e-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="beb6e-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="beb6e-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="beb6e-108">Excel</span><span class="sxs-lookup"><span data-stu-id="beb6e-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="beb6e-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="beb6e-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="beb6e-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="beb6e-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="beb6e-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="beb6e-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="beb6e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="beb6e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="beb6e-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="beb6e-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-114">- TaskPane</span></span><br><span data-ttu-id="beb6e-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-115">
        - Content</span></span><br><span data-ttu-id="beb6e-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="beb6e-116">
        - Custom Functions</span></span><br><span data-ttu-id="beb6e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="beb6e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="beb6e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="beb6e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="beb6e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="beb6e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="beb6e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="beb6e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="beb6e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="beb6e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="beb6e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="beb6e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="beb6e-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-128">
        - BindingEvents</span></span><br><span data-ttu-id="beb6e-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-129">
        - CompressedFile</span></span><br><span data-ttu-id="beb6e-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-130">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-131">
        - File</span></span><br><span data-ttu-id="beb6e-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-132">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-134">
        - Selection</span></span><br><span data-ttu-id="beb6e-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-135">
        - Settings</span></span><br><span data-ttu-id="beb6e-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-136">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-137">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-138">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-140">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-140">Office on Windows</span></span><br><span data-ttu-id="beb6e-141">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-141">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-142">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-142">- TaskPane</span></span><br><span data-ttu-id="beb6e-143">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-143">
        - Content</span></span><br><span data-ttu-id="beb6e-144">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="beb6e-144">
        - Custom Functions</span></span><br><span data-ttu-id="beb6e-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="beb6e-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="beb6e-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="beb6e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="beb6e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="beb6e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="beb6e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="beb6e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="beb6e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="beb6e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="beb6e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="beb6e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="beb6e-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="beb6e-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-158">
        - BindingEvents</span></span><br><span data-ttu-id="beb6e-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-159">
        - CompressedFile</span></span><br><span data-ttu-id="beb6e-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-160">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-161">
        - File</span></span><br><span data-ttu-id="beb6e-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-162">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-164">
        - Selection</span></span><br><span data-ttu-id="beb6e-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-165">
        - Settings</span></span><br><span data-ttu-id="beb6e-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-166">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-167">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-168">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-170">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-170">Office 2019 on Windows</span></span><br><span data-ttu-id="beb6e-171">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="beb6e-172">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-172">- TaskPane</span></span><br><span data-ttu-id="beb6e-173">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-173">
        - Content</span></span><br><span data-ttu-id="beb6e-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="beb6e-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="beb6e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="beb6e-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="beb6e-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="beb6e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="beb6e-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="beb6e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="beb6e-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="beb6e-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="beb6e-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-185">- BindingEvents</span></span><br><span data-ttu-id="beb6e-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-186">
        - CompressedFile</span></span><br><span data-ttu-id="beb6e-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-187">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-188">
        - File</span></span><br><span data-ttu-id="beb6e-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-189">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-191">
        - Selection</span></span><br><span data-ttu-id="beb6e-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-192">
        - Settings</span></span><br><span data-ttu-id="beb6e-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-193">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-194">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-195">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-197">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-197">Office 2016 on Windows</span></span><br><span data-ttu-id="beb6e-198">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="beb6e-199">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-199">- TaskPane</span></span><br><span data-ttu-id="beb6e-200">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-200">
        - Content</span></span></td>
    <td><span data-ttu-id="beb6e-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="beb6e-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="beb6e-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="beb6e-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="beb6e-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-204">- BindingEvents</span></span><br><span data-ttu-id="beb6e-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-205">
        - CompressedFile</span></span><br><span data-ttu-id="beb6e-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-206">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-207">
        - File</span></span><br><span data-ttu-id="beb6e-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-208">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-210">
        - Selection</span></span><br><span data-ttu-id="beb6e-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-211">
        - Settings</span></span><br><span data-ttu-id="beb6e-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-212">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-213">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-214">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-216">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-216">Office 2013 on Windows</span></span><br><span data-ttu-id="beb6e-217">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="beb6e-218">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-218">
        - TaskPane</span></span><br><span data-ttu-id="beb6e-219">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="beb6e-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="beb6e-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="beb6e-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="beb6e-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-222">
        - BindingEvents</span></span><br><span data-ttu-id="beb6e-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-223">
        - CompressedFile</span></span><br><span data-ttu-id="beb6e-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-224">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-225">
        - File</span></span><br><span data-ttu-id="beb6e-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-226">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-228">
        - Selection</span></span><br><span data-ttu-id="beb6e-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-229">
        - Settings</span></span><br><span data-ttu-id="beb6e-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-230">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-231">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-232">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-234">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="beb6e-234">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="beb6e-235">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-235">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="beb6e-236">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-236">- TaskPane</span></span><br><span data-ttu-id="beb6e-237">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-237">
        - Content</span></span></td>
    <td><span data-ttu-id="beb6e-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="beb6e-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="beb6e-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="beb6e-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="beb6e-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="beb6e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="beb6e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="beb6e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="beb6e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="beb6e-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="beb6e-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-249">- BindingEvents</span></span><br><span data-ttu-id="beb6e-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-250">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-251">
        - File</span></span><br><span data-ttu-id="beb6e-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-252">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-254">
        - Selection</span></span><br><span data-ttu-id="beb6e-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-255">
        - Settings</span></span><br><span data-ttu-id="beb6e-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-256">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-257">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-258">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-260">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-260">Office apps on Mac</span></span><br><span data-ttu-id="beb6e-261">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-261">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="beb6e-262">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-262">- TaskPane</span></span><br><span data-ttu-id="beb6e-263">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-263">
        - Content</span></span><br><span data-ttu-id="beb6e-264">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="beb6e-264">
        - Custom Functions</span></span><br><span data-ttu-id="beb6e-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="beb6e-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="beb6e-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="beb6e-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="beb6e-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="beb6e-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="beb6e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="beb6e-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="beb6e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="beb6e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="beb6e-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="beb6e-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="beb6e-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-278">- BindingEvents</span></span><br><span data-ttu-id="beb6e-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-279">
        - CompressedFile</span></span><br><span data-ttu-id="beb6e-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-280">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-281">
        - File</span></span><br><span data-ttu-id="beb6e-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-282">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-284">
        - PdfFile</span></span><br><span data-ttu-id="beb6e-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-285">
        - Selection</span></span><br><span data-ttu-id="beb6e-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-286">
        - Settings</span></span><br><span data-ttu-id="beb6e-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-287">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-288">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-289">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-291">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-291">Office 2019 for Mac</span></span><br><span data-ttu-id="beb6e-292">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="beb6e-293">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-293">- TaskPane</span></span><br><span data-ttu-id="beb6e-294">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-294">
        - Content</span></span><br><span data-ttu-id="beb6e-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="beb6e-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="beb6e-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="beb6e-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="beb6e-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="beb6e-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="beb6e-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="beb6e-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="beb6e-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="beb6e-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="beb6e-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-306">- BindingEvents</span></span><br><span data-ttu-id="beb6e-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-307">
        - CompressedFile</span></span><br><span data-ttu-id="beb6e-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-308">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-309">
        - File</span></span><br><span data-ttu-id="beb6e-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-310">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-312">
        - PdfFile</span></span><br><span data-ttu-id="beb6e-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-313">
        - Selection</span></span><br><span data-ttu-id="beb6e-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-314">
        - Settings</span></span><br><span data-ttu-id="beb6e-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-315">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-316">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-317">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-319">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-319">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="beb6e-320">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="beb6e-321">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-321">- TaskPane</span></span><br><span data-ttu-id="beb6e-322">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-322">
        - Content</span></span></td>
    <td><span data-ttu-id="beb6e-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="beb6e-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="beb6e-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="beb6e-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="beb6e-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-326">- BindingEvents</span></span><br><span data-ttu-id="beb6e-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-327">
        - CompressedFile</span></span><br><span data-ttu-id="beb6e-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-328">
        - DocumentEvents</span></span><br><span data-ttu-id="beb6e-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-329">
        - File</span></span><br><span data-ttu-id="beb6e-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-330">
        - MatrixBindings</span></span><br><span data-ttu-id="beb6e-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-332">
        - PdfFile</span></span><br><span data-ttu-id="beb6e-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-333">
        - Selection</span></span><br><span data-ttu-id="beb6e-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-334">
        - Settings</span></span><br><span data-ttu-id="beb6e-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-335">
        - TableBindings</span></span><br><span data-ttu-id="beb6e-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-336">
        - TableCoercion</span></span><br><span data-ttu-id="beb6e-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-337">
        - TextBindings</span></span><br><span data-ttu-id="beb6e-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="beb6e-339">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="beb6e-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="beb6e-340">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="beb6e-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="beb6e-341">Платформа</span><span class="sxs-lookup"><span data-stu-id="beb6e-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="beb6e-342">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="beb6e-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="beb6e-343">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="beb6e-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="beb6e-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="beb6e-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-345">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="beb6e-345">Office on the web</span></span></td>
    <td><span data-ttu-id="beb6e-346">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="beb6e-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="beb6e-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-348">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-348">Office on Windows</span></span><br><span data-ttu-id="beb6e-349">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-349">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="beb6e-350">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="beb6e-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="beb6e-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-352">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-352">Office for Mac</span></span><br><span data-ttu-id="beb6e-353">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="beb6e-354">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="beb6e-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="beb6e-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="beb6e-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="beb6e-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="beb6e-357">Платформа</span><span class="sxs-lookup"><span data-stu-id="beb6e-357">Platform</span></span></th>
    <th><span data-ttu-id="beb6e-358">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="beb6e-358">Extension points</span></span></th>
    <th><span data-ttu-id="beb6e-359">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="beb6e-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="beb6e-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="beb6e-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-361">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="beb6e-361">Office on the web</span></span><br><span data-ttu-id="beb6e-362">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="beb6e-362">modern</span></span></td>
    <td> <span data-ttu-id="beb6e-363">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-363">- Mail Read</span></span><br><span data-ttu-id="beb6e-364">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-364">
      - Mail Compose</span></span><br><span data-ttu-id="beb6e-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="beb6e-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="beb6e-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="beb6e-373">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-374">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="beb6e-374">Office on the web</span></span><br><span data-ttu-id="beb6e-375">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="beb6e-375">classic</span></span></td>
    <td> <span data-ttu-id="beb6e-376">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-376">- Mail Read</span></span><br><span data-ttu-id="beb6e-377">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-377">
      - Mail Compose</span></span><br><span data-ttu-id="beb6e-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="beb6e-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="beb6e-385">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-386">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-386">Office on Windows</span></span><br><span data-ttu-id="beb6e-387">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-387">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-388">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-388">- Mail Read</span></span><br><span data-ttu-id="beb6e-389">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-389">
      - Mail Compose</span></span><br><span data-ttu-id="beb6e-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="beb6e-391">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="beb6e-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="beb6e-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="beb6e-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="beb6e-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="beb6e-399">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-400">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-400">Office 2019 on Windows</span></span><br><span data-ttu-id="beb6e-401">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-402">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-402">- Mail Read</span></span><br><span data-ttu-id="beb6e-403">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-403">
      - Mail Compose</span></span><br><span data-ttu-id="beb6e-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="beb6e-405">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="beb6e-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="beb6e-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="beb6e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="beb6e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="beb6e-413">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-414">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-414">Office 2016 on Windows</span></span><br><span data-ttu-id="beb6e-415">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-416">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-416">- Mail Read</span></span><br><span data-ttu-id="beb6e-417">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-417">
      - Mail Compose</span></span><br><span data-ttu-id="beb6e-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="beb6e-419">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="beb6e-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="beb6e-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="beb6e-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="beb6e-424">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-425">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-425">Office 2013 on Windows</span></span><br><span data-ttu-id="beb6e-426">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-427">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-427">- Mail Read</span></span><br><span data-ttu-id="beb6e-428">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="beb6e-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="beb6e-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="beb6e-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="beb6e-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="beb6e-433">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-434">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="beb6e-434">Office apps on iOS</span></span><br><span data-ttu-id="beb6e-435">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-435">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-436">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-436">- Mail Read</span></span><br><span data-ttu-id="beb6e-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="beb6e-443">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-444">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-444">Office apps on Mac</span></span><br><span data-ttu-id="beb6e-445">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-445">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-446">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-446">- Mail Read</span></span><br><span data-ttu-id="beb6e-447">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-447">
      - Mail Compose</span></span><br><span data-ttu-id="beb6e-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="beb6e-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="beb6e-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="beb6e-456">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-457">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-457">Office 2019 for Mac</span></span><br><span data-ttu-id="beb6e-458">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-459">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-459">- Mail Read</span></span><br><span data-ttu-id="beb6e-460">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-460">
      - Mail Compose</span></span><br><span data-ttu-id="beb6e-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="beb6e-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="beb6e-468">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-469">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-469">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="beb6e-470">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-471">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-471">- Mail Read</span></span><br><span data-ttu-id="beb6e-472">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-472">
      - Mail Compose</span></span><br><span data-ttu-id="beb6e-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="beb6e-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="beb6e-480">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-481">Office для Android</span><span class="sxs-lookup"><span data-stu-id="beb6e-481">Office apps on Android</span></span><br><span data-ttu-id="beb6e-482">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-482">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-483">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="beb6e-483">- Mail Read</span></span><br><span data-ttu-id="beb6e-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="beb6e-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="beb6e-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="beb6e-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="beb6e-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="beb6e-490">Недоступно</span><span class="sxs-lookup"><span data-stu-id="beb6e-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="beb6e-491">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="beb6e-491">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="beb6e-492">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="beb6e-492">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="beb6e-493">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="beb6e-493">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="beb6e-494">Word</span><span class="sxs-lookup"><span data-stu-id="beb6e-494">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="beb6e-495">Платформа</span><span class="sxs-lookup"><span data-stu-id="beb6e-495">Platform</span></span></th>
    <th><span data-ttu-id="beb6e-496">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="beb6e-496">Extension points</span></span></th>
    <th><span data-ttu-id="beb6e-497">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="beb6e-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="beb6e-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="beb6e-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-499">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="beb6e-499">Office on the web</span></span></td>
    <td> <span data-ttu-id="beb6e-500">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-500">- TaskPane</span></span><br><span data-ttu-id="beb6e-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="beb6e-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="beb6e-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="beb6e-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="beb6e-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="beb6e-508">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-508">- BindingEvents</span></span><br><span data-ttu-id="beb6e-509">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-509">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-510">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-510">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-511">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-511">
         - File</span></span><br><span data-ttu-id="beb6e-512">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-512">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-513">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-513">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-514">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-514">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-515">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-515">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-516">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-516">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-517">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-517">
         - Selection</span></span><br><span data-ttu-id="beb6e-518">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-518">
         - Settings</span></span><br><span data-ttu-id="beb6e-519">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-519">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-520">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-520">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-521">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-521">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-522">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-522">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-523">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-523">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-524">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-524">Office on Windows</span></span><br><span data-ttu-id="beb6e-525">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-525">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-526">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-526">- TaskPane</span></span><br><span data-ttu-id="beb6e-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="beb6e-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="beb6e-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="beb6e-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="beb6e-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="beb6e-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-534">- BindingEvents</span></span><br><span data-ttu-id="beb6e-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-535">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-537">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-538">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-538">
         - File</span></span><br><span data-ttu-id="beb6e-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-540">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-543">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-544">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-544">
         - Selection</span></span><br><span data-ttu-id="beb6e-545">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-545">
         - Settings</span></span><br><span data-ttu-id="beb6e-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-546">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-547">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-548">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-549">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-550">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-551">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-551">Office 2019 on Windows</span></span><br><span data-ttu-id="beb6e-552">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-552">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-553">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-553">- TaskPane</span></span><br><span data-ttu-id="beb6e-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="beb6e-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="beb6e-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="beb6e-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-560">- BindingEvents</span></span><br><span data-ttu-id="beb6e-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-561">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-563">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-564">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-564">
         - File</span></span><br><span data-ttu-id="beb6e-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-566">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-569">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-570">
         - Selection</span></span><br><span data-ttu-id="beb6e-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-571">
         - Settings</span></span><br><span data-ttu-id="beb6e-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-572">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-573">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-574">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-575">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-577">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-577">Office 2016 on Windows</span></span><br><span data-ttu-id="beb6e-578">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-579">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-579">- TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="beb6e-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="beb6e-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="beb6e-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-583">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-583">- BindingEvents</span></span><br><span data-ttu-id="beb6e-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-584">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-585">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-585">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-586">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-586">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-587">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-587">
         - File</span></span><br><span data-ttu-id="beb6e-588">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-588">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-589">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-589">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-590">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-590">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-591">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-591">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-592">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-592">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-593">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-593">
         - Selection</span></span><br><span data-ttu-id="beb6e-594">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-594">
         - Settings</span></span><br><span data-ttu-id="beb6e-595">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-595">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-596">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-596">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-597">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-597">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-598">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-599">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-599">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-600">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-600">Office 2013 on Windows</span></span><br><span data-ttu-id="beb6e-601">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-601">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-602">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-602">- TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="beb6e-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="beb6e-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-605">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-605">- BindingEvents</span></span><br><span data-ttu-id="beb6e-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-606">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-607">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-607">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-608">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-609">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-609">
         - File</span></span><br><span data-ttu-id="beb6e-610">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-610">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-611">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-611">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-612">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-612">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-613">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-613">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-614">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-615">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-615">
         - Selection</span></span><br><span data-ttu-id="beb6e-616">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-616">
         - Settings</span></span><br><span data-ttu-id="beb6e-617">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-617">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-618">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-618">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-619">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-619">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-620">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-620">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-621">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-621">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-622">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="beb6e-622">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="beb6e-623">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-623">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-624">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-624">- TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="beb6e-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="beb6e-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="beb6e-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="beb6e-630">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-630">- BindingEvents</span></span><br><span data-ttu-id="beb6e-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-631">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-632">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-632">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-633">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-634">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-634">
         - File</span></span><br><span data-ttu-id="beb6e-635">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-635">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-636">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-636">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-637">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-637">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-638">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-638">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-639">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-639">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-640">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-640">
         - Selection</span></span><br><span data-ttu-id="beb6e-641">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-641">
         - Settings</span></span><br><span data-ttu-id="beb6e-642">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-642">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-643">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-643">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-644">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-644">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-645">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-645">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-646">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-646">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-647">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-647">Office apps on Mac</span></span><br><span data-ttu-id="beb6e-648">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-648">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-649">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-649">- TaskPane</span></span><br><span data-ttu-id="beb6e-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="beb6e-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="beb6e-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="beb6e-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="beb6e-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="beb6e-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-657">- BindingEvents</span></span><br><span data-ttu-id="beb6e-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-658">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-660">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-661">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-661">
         - File</span></span><br><span data-ttu-id="beb6e-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-663">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-666">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-667">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-667">
         - Selection</span></span><br><span data-ttu-id="beb6e-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-668">
         - Settings</span></span><br><span data-ttu-id="beb6e-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-669">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-670">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-671">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-672">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-674">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-674">Office 2019 for Mac</span></span><br><span data-ttu-id="beb6e-675">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-675">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-676">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-676">- TaskPane</span></span><br><span data-ttu-id="beb6e-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="beb6e-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="beb6e-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="beb6e-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="beb6e-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-683">- BindingEvents</span></span><br><span data-ttu-id="beb6e-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-684">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-686">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-687">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-687">
         - File</span></span><br><span data-ttu-id="beb6e-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-689">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-692">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-693">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-693">
         - Selection</span></span><br><span data-ttu-id="beb6e-694">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-694">
         - Settings</span></span><br><span data-ttu-id="beb6e-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-695">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-696">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-697">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-698">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-700">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-700">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="beb6e-701">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-702">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-702">- TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="beb6e-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="beb6e-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="beb6e-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-706">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-706">- BindingEvents</span></span><br><span data-ttu-id="beb6e-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-707">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-708">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="beb6e-708">
         - CustomXmlParts</span></span><br><span data-ttu-id="beb6e-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-709">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-710">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="beb6e-710">
         - File</span></span><br><span data-ttu-id="beb6e-711">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-711">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-712">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-712">
         - MatrixBindings</span></span><br><span data-ttu-id="beb6e-713">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-713">
         - MatrixCoercion</span></span><br><span data-ttu-id="beb6e-714">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-714">
         - OoxmlCoercion</span></span><br><span data-ttu-id="beb6e-715">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-715">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-716">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-716">
         - Selection</span></span><br><span data-ttu-id="beb6e-717">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="beb6e-717">
         - Settings</span></span><br><span data-ttu-id="beb6e-718">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-718">
         - TableBindings</span></span><br><span data-ttu-id="beb6e-719">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-719">
         - TableCoercion</span></span><br><span data-ttu-id="beb6e-720">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="beb6e-720">
         - TextBindings</span></span><br><span data-ttu-id="beb6e-721">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-721">
         - TextCoercion</span></span><br><span data-ttu-id="beb6e-722">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-722">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="beb6e-723">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="beb6e-723">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="beb6e-724">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="beb6e-724">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="beb6e-725">Платформа</span><span class="sxs-lookup"><span data-stu-id="beb6e-725">Platform</span></span></th>
    <th><span data-ttu-id="beb6e-726">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="beb6e-726">Extension points</span></span></th>
    <th><span data-ttu-id="beb6e-727">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="beb6e-727">API requirement sets</span></span></th>
    <th><span data-ttu-id="beb6e-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="beb6e-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-729">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="beb6e-729">Office on the web</span></span></td>
    <td> <span data-ttu-id="beb6e-730">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-730">- Content</span></span><br><span data-ttu-id="beb6e-731">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-731">
         - TaskPane</span></span><br><span data-ttu-id="beb6e-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="beb6e-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="beb6e-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="beb6e-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-737">- ActiveView</span></span><br><span data-ttu-id="beb6e-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-738">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-739">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-740">
         - File</span></span><br><span data-ttu-id="beb6e-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-741">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-742">
         - Selection</span></span><br><span data-ttu-id="beb6e-743">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-743">
         - Settings</span></span><br><span data-ttu-id="beb6e-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-745">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-745">Office on Windows</span></span><br><span data-ttu-id="beb6e-746">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-746">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-747">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-747">- Content</span></span><br><span data-ttu-id="beb6e-748">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-748">
         - TaskPane</span></span><br><span data-ttu-id="beb6e-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="beb6e-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="beb6e-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="beb6e-754">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-754">- ActiveView</span></span><br><span data-ttu-id="beb6e-755">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-755">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-756">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-756">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-757">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-757">
         - File</span></span><br><span data-ttu-id="beb6e-758">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-758">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-759">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-759">
         - Selection</span></span><br><span data-ttu-id="beb6e-760">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-760">
         - Settings</span></span><br><span data-ttu-id="beb6e-761">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-761">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-762">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-762">Office 2019 on Windows</span></span><br><span data-ttu-id="beb6e-763">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-763">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-764">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-764">- Content</span></span><br><span data-ttu-id="beb6e-765">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-765">
         - TaskPane</span></span><br><span data-ttu-id="beb6e-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-769">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-769">- ActiveView</span></span><br><span data-ttu-id="beb6e-770">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-770">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-771">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-771">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-772">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-772">
         - File</span></span><br><span data-ttu-id="beb6e-773">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-773">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-774">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-774">
         - Selection</span></span><br><span data-ttu-id="beb6e-775">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-775">
         - Settings</span></span><br><span data-ttu-id="beb6e-776">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-776">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-777">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-777">Office 2016 on Windows</span></span><br><span data-ttu-id="beb6e-778">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-778">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-779">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-779">- Content</span></span><br><span data-ttu-id="beb6e-780">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-780">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="beb6e-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="beb6e-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-783">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-783">- ActiveView</span></span><br><span data-ttu-id="beb6e-784">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-784">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-785">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-785">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-786">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-786">
         - File</span></span><br><span data-ttu-id="beb6e-787">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-787">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-788">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-788">
         - Selection</span></span><br><span data-ttu-id="beb6e-789">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-789">
         - Settings</span></span><br><span data-ttu-id="beb6e-790">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-790">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-791">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-791">Office 2013 on Windows</span></span><br><span data-ttu-id="beb6e-792">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-792">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-793">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-793">- Content</span></span><br><span data-ttu-id="beb6e-794">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-794">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="beb6e-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="beb6e-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="beb6e-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-797">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-797">- ActiveView</span></span><br><span data-ttu-id="beb6e-798">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-798">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-799">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-799">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-800">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-800">
         - File</span></span><br><span data-ttu-id="beb6e-801">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-801">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-802">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-802">
         - Selection</span></span><br><span data-ttu-id="beb6e-803">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-803">
         - Settings</span></span><br><span data-ttu-id="beb6e-804">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-804">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-805">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="beb6e-805">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="beb6e-806">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-806">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-807">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-807">- Content</span></span><br><span data-ttu-id="beb6e-808">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-808">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="beb6e-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-812">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-812">- ActiveView</span></span><br><span data-ttu-id="beb6e-813">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-813">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-814">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-814">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-815">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-815">
         - File</span></span><br><span data-ttu-id="beb6e-816">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-816">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-817">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-817">
         - Selection</span></span><br><span data-ttu-id="beb6e-818">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-818">
         - Settings</span></span><br><span data-ttu-id="beb6e-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-819">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-820">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-820">Office apps on Mac</span></span><br><span data-ttu-id="beb6e-821">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="beb6e-821">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="beb6e-822">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-822">- Content</span></span><br><span data-ttu-id="beb6e-823">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-823">
         - TaskPane</span></span><br><span data-ttu-id="beb6e-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="beb6e-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="beb6e-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="beb6e-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-829">- ActiveView</span></span><br><span data-ttu-id="beb6e-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-830">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-831">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-832">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-832">
         - File</span></span><br><span data-ttu-id="beb6e-833">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-833">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-834">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-834">
         - Selection</span></span><br><span data-ttu-id="beb6e-835">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-835">
         - Settings</span></span><br><span data-ttu-id="beb6e-836">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-836">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-837">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-837">Office 2019 for Mac</span></span><br><span data-ttu-id="beb6e-838">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-838">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-839">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-839">- Content</span></span><br><span data-ttu-id="beb6e-840">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-840">
         - TaskPane</span></span><br><span data-ttu-id="beb6e-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-844">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-844">- ActiveView</span></span><br><span data-ttu-id="beb6e-845">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-845">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-846">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-846">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-847">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-847">
         - File</span></span><br><span data-ttu-id="beb6e-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-848">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-849">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-849">
         - Selection</span></span><br><span data-ttu-id="beb6e-850">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-850">
         - Settings</span></span><br><span data-ttu-id="beb6e-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-851">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-852">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-852">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="beb6e-853">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-853">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-854">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-854">- Content</span></span><br><span data-ttu-id="beb6e-855">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-855">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="beb6e-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="beb6e-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-858">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="beb6e-858">- ActiveView</span></span><br><span data-ttu-id="beb6e-859">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-859">
         - CompressedFile</span></span><br><span data-ttu-id="beb6e-860">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-860">
         - DocumentEvents</span></span><br><span data-ttu-id="beb6e-861">
         - File</span><span class="sxs-lookup"><span data-stu-id="beb6e-861">
         - File</span></span><br><span data-ttu-id="beb6e-862">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="beb6e-862">
         - PdfFile</span></span><br><span data-ttu-id="beb6e-863">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-863">
         - Selection</span></span><br><span data-ttu-id="beb6e-864">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-864">
         - Settings</span></span><br><span data-ttu-id="beb6e-865">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-865">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="beb6e-866">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="beb6e-866">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="beb6e-867">OneNote</span><span class="sxs-lookup"><span data-stu-id="beb6e-867">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="beb6e-868">Платформа</span><span class="sxs-lookup"><span data-stu-id="beb6e-868">Platform</span></span></th>
    <th><span data-ttu-id="beb6e-869">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="beb6e-869">Extension points</span></span></th>
    <th><span data-ttu-id="beb6e-870">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="beb6e-870">API requirement sets</span></span></th>
    <th><span data-ttu-id="beb6e-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="beb6e-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-872">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="beb6e-872">Office on the web</span></span></td>
    <td> <span data-ttu-id="beb6e-873">- Контент</span><span class="sxs-lookup"><span data-stu-id="beb6e-873">- Content</span></span><br><span data-ttu-id="beb6e-874">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-874">
         - TaskPane</span></span><br><span data-ttu-id="beb6e-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="beb6e-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="beb6e-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="beb6e-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-879">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="beb6e-879">- DocumentEvents</span></span><br><span data-ttu-id="beb6e-880">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-880">
         - HtmlCoercion</span></span><br><span data-ttu-id="beb6e-881">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="beb6e-881">
         - Settings</span></span><br><span data-ttu-id="beb6e-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="beb6e-883">Project</span><span class="sxs-lookup"><span data-stu-id="beb6e-883">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="beb6e-884">Платформа</span><span class="sxs-lookup"><span data-stu-id="beb6e-884">Platform</span></span></th>
    <th><span data-ttu-id="beb6e-885">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="beb6e-885">Extension points</span></span></th>
    <th><span data-ttu-id="beb6e-886">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="beb6e-886">API requirement sets</span></span></th>
    <th><span data-ttu-id="beb6e-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="beb6e-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-888">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-888">Office 2019 on Windows</span></span><br><span data-ttu-id="beb6e-889">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-889">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-890">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-890">- TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-892">- Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-892">- Selection</span></span><br><span data-ttu-id="beb6e-893">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-893">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-894">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-894">Office 2016 on Windows</span></span><br><span data-ttu-id="beb6e-895">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-895">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-896">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-896">- TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-898">- Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-898">- Selection</span></span><br><span data-ttu-id="beb6e-899">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-899">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="beb6e-900">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="beb6e-900">Office 2013 on Windows</span></span><br><span data-ttu-id="beb6e-901">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="beb6e-901">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="beb6e-902">- Область задач</span><span class="sxs-lookup"><span data-stu-id="beb6e-902">- TaskPane</span></span></td>
    <td> <span data-ttu-id="beb6e-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="beb6e-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="beb6e-904">- Selection</span><span class="sxs-lookup"><span data-stu-id="beb6e-904">- Selection</span></span><br><span data-ttu-id="beb6e-905">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="beb6e-905">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="beb6e-906">См. также</span><span class="sxs-lookup"><span data-stu-id="beb6e-906">See also</span></span>

- [<span data-ttu-id="beb6e-907">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="beb6e-907">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="beb6e-908">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="beb6e-908">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="beb6e-909">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="beb6e-909">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="beb6e-910">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="beb6e-910">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="beb6e-911">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="beb6e-911">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="beb6e-912">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="beb6e-912">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="beb6e-913">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="beb6e-913">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="beb6e-914">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="beb6e-914">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="beb6e-915">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="beb6e-915">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="beb6e-916">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="beb6e-916">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="beb6e-917">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="beb6e-917">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
