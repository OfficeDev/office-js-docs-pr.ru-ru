---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 19f2fa7f744345823c2700b04524ec20705035a8
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952371"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="0ad7e-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0ad7e-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="0ad7e-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="0ad7e-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="0ad7e-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="0ad7e-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="0ad7e-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="0ad7e-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="0ad7e-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="0ad7e-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="0ad7e-108">Excel</span><span class="sxs-lookup"><span data-stu-id="0ad7e-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="0ad7e-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="0ad7e-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="0ad7e-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0ad7e-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="0ad7e-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0ad7e-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="0ad7e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="0ad7e-113">Office Online</span></span></td>
    <td> <span data-ttu-id="0ad7e-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-114">- TaskPane</span></span><br><span data-ttu-id="0ad7e-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-115">
        - Content</span></span><br><span data-ttu-id="0ad7e-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-116">
        -Custom Functions</span></span><br><span data-ttu-id="0ad7e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="0ad7e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="0ad7e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0ad7e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0ad7e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0ad7e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0ad7e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0ad7e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="0ad7e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0ad7e-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-128">
        - BindingEvents</span></span><br><span data-ttu-id="0ad7e-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-129">
        - CompressedFile</span></span><br><span data-ttu-id="0ad7e-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-130">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-131">
        - File</span></span><br><span data-ttu-id="0ad7e-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-132">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-134">
        - Selection</span></span><br><span data-ttu-id="0ad7e-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-135">
        - Settings</span></span><br><span data-ttu-id="0ad7e-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-136">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-137">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-138">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-140">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-140">Office apps on Windows</span></span><br><span data-ttu-id="0ad7e-141">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-142">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-142">- TaskPane</span></span><br><span data-ttu-id="0ad7e-143">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-143">
        - Content</span></span><br><span data-ttu-id="0ad7e-144">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-144">
        -Custom Functions</span></span><br><span data-ttu-id="0ad7e-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="0ad7e-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="0ad7e-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0ad7e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0ad7e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0ad7e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0ad7e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0ad7e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="0ad7e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0ad7e-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-156">
        - BindingEvents</span></span><br><span data-ttu-id="0ad7e-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-157">
        - CompressedFile</span></span><br><span data-ttu-id="0ad7e-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-158">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-159">
        - File</span></span><br><span data-ttu-id="0ad7e-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-160">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-162">
        - Selection</span></span><br><span data-ttu-id="0ad7e-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-163">
        - Settings</span></span><br><span data-ttu-id="0ad7e-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-164">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-165">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-166">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-168">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-168">Office 2019 for Windows</span></span><br><span data-ttu-id="0ad7e-169">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="0ad7e-170">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-170">- TaskPane</span></span><br><span data-ttu-id="0ad7e-171">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-171">
        - Content</span></span><br><span data-ttu-id="0ad7e-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0ad7e-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0ad7e-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0ad7e-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0ad7e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0ad7e-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0ad7e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0ad7e-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-182">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-183">
        - CompressedFile</span></span><br><span data-ttu-id="0ad7e-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-184">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-185">
        - File</span></span><br><span data-ttu-id="0ad7e-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-186">
        - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-187">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-189">
        - Selection</span></span><br><span data-ttu-id="0ad7e-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-190">
        - Settings</span></span><br><span data-ttu-id="0ad7e-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-191">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-192">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-193">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-195">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-195">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="0ad7e-196">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="0ad7e-197">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-197">- TaskPane</span></span><br><span data-ttu-id="0ad7e-198">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-198">
        - Content</span></span></td>
    <td><span data-ttu-id="0ad7e-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="0ad7e-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-201">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-202">
        - CompressedFile</span></span><br><span data-ttu-id="0ad7e-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-203">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-204">
        - File</span></span><br><span data-ttu-id="0ad7e-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-205">
        - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-206">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-208">
        - Selection</span></span><br><span data-ttu-id="0ad7e-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-209">
        - Settings</span></span><br><span data-ttu-id="0ad7e-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-210">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-211">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-212">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-214">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-214">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="0ad7e-215">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="0ad7e-216">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-216">
        - TaskPane</span></span><br><span data-ttu-id="0ad7e-217">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="0ad7e-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="0ad7e-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-219">
        - BindingEvents</span></span><br><span data-ttu-id="0ad7e-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-220">
        - CompressedFile</span></span><br><span data-ttu-id="0ad7e-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-221">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-222">
        - File</span></span><br><span data-ttu-id="0ad7e-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-223">
        - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-224">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-226">
        - Selection</span></span><br><span data-ttu-id="0ad7e-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-227">
        - Settings</span></span><br><span data-ttu-id="0ad7e-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-228">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-229">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-230">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-232">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="0ad7e-232">Office for iPad</span></span><br><span data-ttu-id="0ad7e-233">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="0ad7e-234">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-234">- TaskPane</span></span><br><span data-ttu-id="0ad7e-235">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-235">
        - Content</span></span><br><span data-ttu-id="0ad7e-236">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-236">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="0ad7e-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0ad7e-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0ad7e-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0ad7e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0ad7e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0ad7e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="0ad7e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0ad7e-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-247">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-248">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-249">
        - File</span></span><br><span data-ttu-id="0ad7e-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-250">
        - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-251">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-253">
        - Selection</span></span><br><span data-ttu-id="0ad7e-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-254">
        - Settings</span></span><br><span data-ttu-id="0ad7e-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-255">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-256">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-257">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-259">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-259">Office for Mac</span></span><br><span data-ttu-id="0ad7e-260">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="0ad7e-261">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-261">- TaskPane</span></span><br><span data-ttu-id="0ad7e-262">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-262">
        - Content</span></span><br><span data-ttu-id="0ad7e-263">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-263">
        -Custom Functions</span></span><br><span data-ttu-id="0ad7e-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0ad7e-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0ad7e-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0ad7e-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0ad7e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0ad7e-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0ad7e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="0ad7e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0ad7e-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-275">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-276">
        - CompressedFile</span></span><br><span data-ttu-id="0ad7e-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-277">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-278">
        - File</span></span><br><span data-ttu-id="0ad7e-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-279">
        - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-280">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-282">
        - PdfFile</span></span><br><span data-ttu-id="0ad7e-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-283">
        - Selection</span></span><br><span data-ttu-id="0ad7e-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-284">
        - Settings</span></span><br><span data-ttu-id="0ad7e-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-285">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-286">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-287">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-289">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-289">Office 2019 for Mac</span></span><br><span data-ttu-id="0ad7e-290">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="0ad7e-291">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-291">- TaskPane</span></span><br><span data-ttu-id="0ad7e-292">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-292">
        - Content</span></span><br><span data-ttu-id="0ad7e-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0ad7e-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0ad7e-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0ad7e-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0ad7e-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0ad7e-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0ad7e-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0ad7e-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-303">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-304">
        - CompressedFile</span></span><br><span data-ttu-id="0ad7e-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-305">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-306">
        - File</span></span><br><span data-ttu-id="0ad7e-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-307">
        - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-308">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-310">
        - PdfFile</span></span><br><span data-ttu-id="0ad7e-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-311">
        - Selection</span></span><br><span data-ttu-id="0ad7e-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-312">
        - Settings</span></span><br><span data-ttu-id="0ad7e-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-313">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-314">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-315">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-317">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-317">Office 2016 for Mac</span></span><br><span data-ttu-id="0ad7e-318">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="0ad7e-319">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-319">- TaskPane</span></span><br><span data-ttu-id="0ad7e-320">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-320">
        - Content</span></span></td>
    <td><span data-ttu-id="0ad7e-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="0ad7e-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-323">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-324">
        - CompressedFile</span></span><br><span data-ttu-id="0ad7e-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-325">
        - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-326">
        - File</span></span><br><span data-ttu-id="0ad7e-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-327">
        - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-328">
        - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-330">
        - PdfFile</span></span><br><span data-ttu-id="0ad7e-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-331">
        - Selection</span></span><br><span data-ttu-id="0ad7e-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-332">
        - Settings</span></span><br><span data-ttu-id="0ad7e-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-333">
        - TableBindings</span></span><br><span data-ttu-id="0ad7e-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-334">
        - TableCoercion</span></span><br><span data-ttu-id="0ad7e-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-335">
        - TextBindings</span></span><br><span data-ttu-id="0ad7e-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="0ad7e-337">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="0ad7e-338">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="0ad7e-339">Платформа</span><span class="sxs-lookup"><span data-stu-id="0ad7e-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="0ad7e-340">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0ad7e-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="0ad7e-341">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0ad7e-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="0ad7e-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="0ad7e-343">Office Online</span></span></td>
    <td><span data-ttu-id="0ad7e-344">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-344">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="0ad7e-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-346">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-346">Office apps on Windows</span></span><br><span data-ttu-id="0ad7e-347">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="0ad7e-348">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-348">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="0ad7e-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-350">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="0ad7e-350">Office for iPad</span></span><br><span data-ttu-id="0ad7e-351">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="0ad7e-352">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-352">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="0ad7e-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-354">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-354">Office for Mac</span></span><br><span data-ttu-id="0ad7e-355">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="0ad7e-356">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="0ad7e-356">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="0ad7e-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="0ad7e-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="0ad7e-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0ad7e-359">Платформа</span><span class="sxs-lookup"><span data-stu-id="0ad7e-359">Platform</span></span></th>
    <th><span data-ttu-id="0ad7e-360">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0ad7e-360">Extension points</span></span></th>
    <th><span data-ttu-id="0ad7e-361">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0ad7e-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="0ad7e-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="0ad7e-363">Office Online</span></span></td>
    <td> <span data-ttu-id="0ad7e-364">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-364">- Mail Read</span></span><br><span data-ttu-id="0ad7e-365">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-365">
      - Mail Compose</span></span><br><span data-ttu-id="0ad7e-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0ad7e-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0ad7e-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0ad7e-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0ad7e-374">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-375">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-375">Office apps on Windows</span></span><br><span data-ttu-id="0ad7e-376">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-377">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-377">- Mail Read</span></span><br><span data-ttu-id="0ad7e-378">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-378">
      - Mail Compose</span></span><br><span data-ttu-id="0ad7e-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0ad7e-380">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="0ad7e-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0ad7e-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0ad7e-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0ad7e-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0ad7e-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0ad7e-388">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-389">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-389">Office 2019 for Windows</span></span><br><span data-ttu-id="0ad7e-390">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-391">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-391">- Mail Read</span></span><br><span data-ttu-id="0ad7e-392">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-392">
      - Mail Compose</span></span><br><span data-ttu-id="0ad7e-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0ad7e-394">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="0ad7e-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0ad7e-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0ad7e-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0ad7e-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0ad7e-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0ad7e-402">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-403">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-403">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="0ad7e-404">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-405">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-405">- Mail Read</span></span><br><span data-ttu-id="0ad7e-406">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-406">
      - Mail Compose</span></span><br><span data-ttu-id="0ad7e-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0ad7e-408">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="0ad7e-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0ad7e-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="0ad7e-413">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-414">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-414">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="0ad7e-415">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-416">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-416">- Mail Read</span></span><br><span data-ttu-id="0ad7e-417">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="0ad7e-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="0ad7e-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="0ad7e-422">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-423">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="0ad7e-423">Office for iOS</span></span><br><span data-ttu-id="0ad7e-424">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-425">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-425">- Mail Read</span></span><br><span data-ttu-id="0ad7e-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0ad7e-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="0ad7e-432">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-433">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-433">Office for Mac</span></span><br><span data-ttu-id="0ad7e-434">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-435">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-435">- Mail Read</span></span><br><span data-ttu-id="0ad7e-436">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-436">
      - Mail Compose</span></span><br><span data-ttu-id="0ad7e-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0ad7e-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0ad7e-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0ad7e-444">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-445">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-445">Office 2019 for Mac</span></span><br><span data-ttu-id="0ad7e-446">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-447">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-447">- Mail Read</span></span><br><span data-ttu-id="0ad7e-448">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-448">
      - Mail Compose</span></span><br><span data-ttu-id="0ad7e-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0ad7e-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0ad7e-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0ad7e-456">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-457">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-457">Office 2016 for Mac</span></span><br><span data-ttu-id="0ad7e-458">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-459">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-459">- Mail Read</span></span><br><span data-ttu-id="0ad7e-460">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-460">
      - Mail Compose</span></span><br><span data-ttu-id="0ad7e-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0ad7e-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0ad7e-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0ad7e-468">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-469">Office для Android</span><span class="sxs-lookup"><span data-stu-id="0ad7e-469">Office for Android</span></span><br><span data-ttu-id="0ad7e-470">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-470">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-471">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0ad7e-471">- Mail Read</span></span><br><span data-ttu-id="0ad7e-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0ad7e-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0ad7e-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0ad7e-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0ad7e-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="0ad7e-478">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0ad7e-478">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="0ad7e-479">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-479">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="0ad7e-480">Word</span><span class="sxs-lookup"><span data-stu-id="0ad7e-480">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0ad7e-481">Платформа</span><span class="sxs-lookup"><span data-stu-id="0ad7e-481">Platform</span></span></th>
    <th><span data-ttu-id="0ad7e-482">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0ad7e-482">Extension points</span></span></th>
    <th><span data-ttu-id="0ad7e-483">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0ad7e-483">API requirement sets</span></span></th>
    <th><span data-ttu-id="0ad7e-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-485">Office Online</span><span class="sxs-lookup"><span data-stu-id="0ad7e-485">Office Online</span></span></td>
    <td> <span data-ttu-id="0ad7e-486">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-486">- TaskPane</span></span><br><span data-ttu-id="0ad7e-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-492">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-492">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-493">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-493">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-494">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-494">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-495">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-495">
         - File</span></span><br><span data-ttu-id="0ad7e-496">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-496">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-497">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-497">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-498">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-498">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-499">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-499">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-500">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-500">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-501">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-501">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-502">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-502">
         - Selection</span></span><br><span data-ttu-id="0ad7e-503">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-503">
         - Settings</span></span><br><span data-ttu-id="0ad7e-504">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-504">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-505">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-505">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-506">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-506">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-507">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-507">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-508">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-508">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-509">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-509">Office apps on Windows</span></span><br><span data-ttu-id="0ad7e-510">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-510">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-511">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-511">- TaskPane</span></span><br><span data-ttu-id="0ad7e-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-517">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-517">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-518">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-518">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-519">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-519">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-520">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-520">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-521">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-521">
         - File</span></span><br><span data-ttu-id="0ad7e-522">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-522">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-523">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-523">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-524">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-524">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-525">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-525">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-526">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-526">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-527">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-527">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-528">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-528">
         - Selection</span></span><br><span data-ttu-id="0ad7e-529">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-529">
         - Settings</span></span><br><span data-ttu-id="0ad7e-530">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-530">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-531">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-531">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-532">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-532">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-533">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-533">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-534">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-534">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-535">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-535">Office 2019 for Windows</span></span><br><span data-ttu-id="0ad7e-536">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-536">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-537">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-537">- TaskPane</span></span><br><span data-ttu-id="0ad7e-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-543">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-543">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-544">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-544">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-545">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-545">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-546">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-546">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-547">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-547">
         - File</span></span><br><span data-ttu-id="0ad7e-548">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-548">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-549">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-549">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-550">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-550">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-551">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-551">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-552">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-552">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-553">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-553">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-554">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-554">
         - Selection</span></span><br><span data-ttu-id="0ad7e-555">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-555">
         - Settings</span></span><br><span data-ttu-id="0ad7e-556">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-556">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-557">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-557">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-558">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-558">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-559">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-559">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-560">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-560">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-561">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-561">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="0ad7e-562">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-562">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-563">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-563">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="0ad7e-566">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-566">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-567">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-567">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-568">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-568">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-569">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-570">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-570">
         - File</span></span><br><span data-ttu-id="0ad7e-571">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-571">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-572">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-572">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-573">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-573">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-574">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-574">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-575">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-575">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-576">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-576">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-577">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-577">
         - Selection</span></span><br><span data-ttu-id="0ad7e-578">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-578">
         - Settings</span></span><br><span data-ttu-id="0ad7e-579">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-579">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-580">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-580">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-581">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-581">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-582">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-582">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-583">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-583">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-584">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-584">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="0ad7e-585">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-585">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-586">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-586">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0ad7e-588">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-588">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-589">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-589">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-590">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-590">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-591">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-592">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-592">
         - File</span></span><br><span data-ttu-id="0ad7e-593">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-593">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-594">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-595">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-595">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-596">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-596">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-597">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-597">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-598">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-599">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-599">
         - Selection</span></span><br><span data-ttu-id="0ad7e-600">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-600">
         - Settings</span></span><br><span data-ttu-id="0ad7e-601">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-601">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-602">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-602">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-603">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-603">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-604">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-604">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-605">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-605">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-606">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="0ad7e-606">Office for iPad</span></span><br><span data-ttu-id="0ad7e-607">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-607">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-608">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-608">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0ad7e-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0ad7e-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-613">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-614">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-616">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-617">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-617">
         - File</span></span><br><span data-ttu-id="0ad7e-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-619">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-619">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-620">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-623">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-624">
         - Selection</span></span><br><span data-ttu-id="0ad7e-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-625">
         - Settings</span></span><br><span data-ttu-id="0ad7e-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-626">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-627">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-628">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-629">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-631">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-631">Office for Mac</span></span><br><span data-ttu-id="0ad7e-632">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-632">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-633">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-633">- TaskPane</span></span><br><span data-ttu-id="0ad7e-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0ad7e-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0ad7e-639">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-639">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-640">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-640">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-641">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-641">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-642">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-642">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-643">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-643">
         - File</span></span><br><span data-ttu-id="0ad7e-644">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-644">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-645">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-645">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-646">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-646">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-647">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-647">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-648">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-648">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-649">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-650">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-650">
         - Selection</span></span><br><span data-ttu-id="0ad7e-651">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-651">
         - Settings</span></span><br><span data-ttu-id="0ad7e-652">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-652">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-653">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-653">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-654">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-654">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-655">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-656">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-656">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-657">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-657">Office 2019 for Mac</span></span><br><span data-ttu-id="0ad7e-658">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-658">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-659">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-659">- TaskPane</span></span><br><span data-ttu-id="0ad7e-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0ad7e-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0ad7e-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0ad7e-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0ad7e-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-665">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-666">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-668">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-669">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-669">
         - File</span></span><br><span data-ttu-id="0ad7e-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-671">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-671">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-672">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-672">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-673">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-673">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-674">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-674">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-675">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-675">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-676">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-676">
         - Selection</span></span><br><span data-ttu-id="0ad7e-677">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-677">
         - Settings</span></span><br><span data-ttu-id="0ad7e-678">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-678">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-679">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-679">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-680">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-680">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-681">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-681">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-682">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-682">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-683">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-683">Office 2016 for Mac</span></span><br><span data-ttu-id="0ad7e-684">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-684">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-685">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="0ad7e-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-688">- BindingEvents</span></span><br><span data-ttu-id="0ad7e-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-689">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0ad7e-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="0ad7e-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-691">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-692">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0ad7e-692">
         - File</span></span><br><span data-ttu-id="0ad7e-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-694">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-694">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-695">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-695">
         - MatrixBindings</span></span><br><span data-ttu-id="0ad7e-696">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-696">
         - MatrixCoercion</span></span><br><span data-ttu-id="0ad7e-697">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-697">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0ad7e-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-698">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-699">
         - Selection</span></span><br><span data-ttu-id="0ad7e-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-700">
         - Settings</span></span><br><span data-ttu-id="0ad7e-701">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-701">
         - TableBindings</span></span><br><span data-ttu-id="0ad7e-702">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-702">
         - TableCoercion</span></span><br><span data-ttu-id="0ad7e-703">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0ad7e-703">
         - TextBindings</span></span><br><span data-ttu-id="0ad7e-704">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-704">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-705">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-705">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="0ad7e-706">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-706">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="0ad7e-707">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0ad7e-707">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0ad7e-708">Платформа</span><span class="sxs-lookup"><span data-stu-id="0ad7e-708">Platform</span></span></th>
    <th><span data-ttu-id="0ad7e-709">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0ad7e-709">Extension points</span></span></th>
    <th><span data-ttu-id="0ad7e-710">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0ad7e-710">API requirement sets</span></span></th>
    <th><span data-ttu-id="0ad7e-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-712">Office Online</span><span class="sxs-lookup"><span data-stu-id="0ad7e-712">Office Online</span></span></td>
    <td> <span data-ttu-id="0ad7e-713">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-713">- Content</span></span><br><span data-ttu-id="0ad7e-714">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-714">
         - TaskPane</span></span><br><span data-ttu-id="0ad7e-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-717">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-717">- ActiveView</span></span><br><span data-ttu-id="0ad7e-718">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-718">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-719">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-719">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-720">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-720">
         - File</span></span><br><span data-ttu-id="0ad7e-721">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-721">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-722">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-723">
         - Selection</span></span><br><span data-ttu-id="0ad7e-724">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-724">
         - Settings</span></span><br><span data-ttu-id="0ad7e-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-725">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-726">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-726">Office apps on Windows</span></span><br><span data-ttu-id="0ad7e-727">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-727">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-728">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-728">- Content</span></span><br><span data-ttu-id="0ad7e-729">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-729">
         - TaskPane</span></span><br><span data-ttu-id="0ad7e-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-732">- ActiveView</span></span><br><span data-ttu-id="0ad7e-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-733">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-734">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-735">
         - File</span></span><br><span data-ttu-id="0ad7e-736">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-736">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-737">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-738">
         - Selection</span></span><br><span data-ttu-id="0ad7e-739">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-739">
         - Settings</span></span><br><span data-ttu-id="0ad7e-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-740">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-741">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-741">Office 2019 for Windows</span></span><br><span data-ttu-id="0ad7e-742">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-742">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-743">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-743">- Content</span></span><br><span data-ttu-id="0ad7e-744">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-744">
         - TaskPane</span></span><br><span data-ttu-id="0ad7e-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-747">- ActiveView</span></span><br><span data-ttu-id="0ad7e-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-748">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-749">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-750">
         - File</span></span><br><span data-ttu-id="0ad7e-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-751">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-752">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-753">
         - Selection</span></span><br><span data-ttu-id="0ad7e-754">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-754">
         - Settings</span></span><br><span data-ttu-id="0ad7e-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-756">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-756">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="0ad7e-757">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-757">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-758">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-758">- Content</span></span><br><span data-ttu-id="0ad7e-759">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-759">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0ad7e-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-761">- ActiveView</span></span><br><span data-ttu-id="0ad7e-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-762">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-763">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-764">
         - File</span></span><br><span data-ttu-id="0ad7e-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-765">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-766">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-767">
         - Selection</span></span><br><span data-ttu-id="0ad7e-768">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-768">
         - Settings</span></span><br><span data-ttu-id="0ad7e-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-770">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-770">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="0ad7e-771">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-772">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-772">- Content</span></span><br><span data-ttu-id="0ad7e-773">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-773">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="0ad7e-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0ad7e-775">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-775">- ActiveView</span></span><br><span data-ttu-id="0ad7e-776">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-776">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-777">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-777">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-778">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-778">
         - File</span></span><br><span data-ttu-id="0ad7e-779">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-779">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-780">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-780">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-781">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-781">
         - Selection</span></span><br><span data-ttu-id="0ad7e-782">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-782">
         - Settings</span></span><br><span data-ttu-id="0ad7e-783">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-783">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-784">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="0ad7e-784">Office for iPad</span></span><br><span data-ttu-id="0ad7e-785">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-785">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-786">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-786">- Content</span></span><br><span data-ttu-id="0ad7e-787">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-787">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-789">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-789">- ActiveView</span></span><br><span data-ttu-id="0ad7e-790">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-790">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-791">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-791">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-792">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-792">
         - File</span></span><br><span data-ttu-id="0ad7e-793">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-793">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-794">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-794">
         - Selection</span></span><br><span data-ttu-id="0ad7e-795">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-795">
         - Settings</span></span><br><span data-ttu-id="0ad7e-796">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-796">
         - TextCoercion</span></span><br><span data-ttu-id="0ad7e-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-797">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-798">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-798">Office for Mac</span></span><br><span data-ttu-id="0ad7e-799">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-799">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="0ad7e-800">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-800">- Content</span></span><br><span data-ttu-id="0ad7e-801">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-801">
         - TaskPane</span></span><br><span data-ttu-id="0ad7e-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-804">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-804">- ActiveView</span></span><br><span data-ttu-id="0ad7e-805">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-805">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-806">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-806">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-807">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-807">
         - File</span></span><br><span data-ttu-id="0ad7e-808">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-808">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-809">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-810">
         - Selection</span></span><br><span data-ttu-id="0ad7e-811">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-811">
         - Settings</span></span><br><span data-ttu-id="0ad7e-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-813">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-813">Office 2019 for Mac</span></span><br><span data-ttu-id="0ad7e-814">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-814">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-815">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-815">- Content</span></span><br><span data-ttu-id="0ad7e-816">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-816">
         - TaskPane</span></span><br><span data-ttu-id="0ad7e-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-819">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-819">- ActiveView</span></span><br><span data-ttu-id="0ad7e-820">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-820">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-821">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-821">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-822">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-822">
         - File</span></span><br><span data-ttu-id="0ad7e-823">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-823">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-824">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-825">
         - Selection</span></span><br><span data-ttu-id="0ad7e-826">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-826">
         - Settings</span></span><br><span data-ttu-id="0ad7e-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-828">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-828">Office 2016 for Mac</span></span><br><span data-ttu-id="0ad7e-829">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-829">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-830">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-830">- Content</span></span><br><span data-ttu-id="0ad7e-831">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-831">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0ad7e-833">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0ad7e-833">- ActiveView</span></span><br><span data-ttu-id="0ad7e-834">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-834">
         - CompressedFile</span></span><br><span data-ttu-id="0ad7e-835">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-835">
         - DocumentEvents</span></span><br><span data-ttu-id="0ad7e-836">
         - File</span><span class="sxs-lookup"><span data-stu-id="0ad7e-836">
         - File</span></span><br><span data-ttu-id="0ad7e-837">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-837">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-838">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0ad7e-838">
         - PdfFile</span></span><br><span data-ttu-id="0ad7e-839">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-839">
         - Selection</span></span><br><span data-ttu-id="0ad7e-840">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-840">
         - Settings</span></span><br><span data-ttu-id="0ad7e-841">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-841">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="0ad7e-842">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="0ad7e-842">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="0ad7e-843">OneNote</span><span class="sxs-lookup"><span data-stu-id="0ad7e-843">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0ad7e-844">Платформа</span><span class="sxs-lookup"><span data-stu-id="0ad7e-844">Platform</span></span></th>
    <th><span data-ttu-id="0ad7e-845">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0ad7e-845">Extension points</span></span></th>
    <th><span data-ttu-id="0ad7e-846">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0ad7e-846">API requirement sets</span></span></th>
    <th><span data-ttu-id="0ad7e-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-848">Office Online</span><span class="sxs-lookup"><span data-stu-id="0ad7e-848">Office Online</span></span></td>
    <td> <span data-ttu-id="0ad7e-849">- Контент</span><span class="sxs-lookup"><span data-stu-id="0ad7e-849">- Content</span></span><br><span data-ttu-id="0ad7e-850">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-850">
         - TaskPane</span></span><br><span data-ttu-id="0ad7e-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="0ad7e-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-854">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0ad7e-854">- DocumentEvents</span></span><br><span data-ttu-id="0ad7e-855">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-855">
         - HtmlCoercion</span></span><br><span data-ttu-id="0ad7e-856">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-856">
         - ImageCoercion</span></span><br><span data-ttu-id="0ad7e-857">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0ad7e-857">
         - Settings</span></span><br><span data-ttu-id="0ad7e-858">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-858">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="0ad7e-859">Project</span><span class="sxs-lookup"><span data-stu-id="0ad7e-859">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0ad7e-860">Платформа</span><span class="sxs-lookup"><span data-stu-id="0ad7e-860">Platform</span></span></th>
    <th><span data-ttu-id="0ad7e-861">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0ad7e-861">Extension points</span></span></th>
    <th><span data-ttu-id="0ad7e-862">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0ad7e-862">API requirement sets</span></span></th>
    <th><span data-ttu-id="0ad7e-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-864">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-864">Office 2019 for Windows</span></span><br><span data-ttu-id="0ad7e-865">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-866">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-866">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-868">- Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-868">- Selection</span></span><br><span data-ttu-id="0ad7e-869">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-869">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-870">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-870">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="0ad7e-871">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-871">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-872">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-872">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-874">- Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-874">- Selection</span></span><br><span data-ttu-id="0ad7e-875">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-875">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0ad7e-876">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0ad7e-876">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="0ad7e-877">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-877">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="0ad7e-878">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0ad7e-878">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0ad7e-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0ad7e-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0ad7e-880">- Selection</span><span class="sxs-lookup"><span data-stu-id="0ad7e-880">- Selection</span></span><br><span data-ttu-id="0ad7e-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0ad7e-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="0ad7e-882">См. также</span><span class="sxs-lookup"><span data-stu-id="0ad7e-882">See also</span></span>

- [<span data-ttu-id="0ad7e-883">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0ad7e-883">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="0ad7e-884">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="0ad7e-884">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="0ad7e-885">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="0ad7e-885">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="0ad7e-886">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="0ad7e-886">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="0ad7e-887">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="0ad7e-887">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="0ad7e-888">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="0ad7e-888">Update history for Office 365 ProPlus releases</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="0ad7e-889">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="0ad7e-889">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="0ad7e-890">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="0ad7e-890">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="0ad7e-891">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-891">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="0ad7e-892">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="0ad7e-892">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="0ad7e-893">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="0ad7e-893">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
