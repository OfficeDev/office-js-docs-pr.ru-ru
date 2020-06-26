---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 06/23/2020
localization_priority: Priority
ms.openlocfilehash: 979c873b1c5f2d1d7847414f037d5c75737aa33d
ms.sourcegitcommit: a4873c3525c7d30ef551545d27eb2c0a16b4eb50
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44888161"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="fe20a-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="fe20a-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="fe20a-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="fe20a-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="fe20a-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="fe20a-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="fe20a-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="fe20a-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="fe20a-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="fe20a-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="fe20a-108">Excel</span><span class="sxs-lookup"><span data-stu-id="fe20a-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="fe20a-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="fe20a-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="fe20a-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="fe20a-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="fe20a-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="fe20a-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="fe20a-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="fe20a-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fe20a-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="fe20a-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-114">- TaskPane</span></span><br><span data-ttu-id="fe20a-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-115">
        - Content</span></span><br><span data-ttu-id="fe20a-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="fe20a-116">
        - Custom Functions</span></span><br><span data-ttu-id="fe20a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="fe20a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fe20a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fe20a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fe20a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fe20a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fe20a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fe20a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fe20a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fe20a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fe20a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="fe20a-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="fe20a-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="fe20a-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="fe20a-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="fe20a-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-131">
        - BindingEvents</span></span><br><span data-ttu-id="fe20a-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-132">
        - CompressedFile</span></span><br><span data-ttu-id="fe20a-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-133">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-134">
        - File</span></span><br><span data-ttu-id="fe20a-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-135">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-137">
        - Selection</span></span><br><span data-ttu-id="fe20a-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-138">
        - Settings</span></span><br><span data-ttu-id="fe20a-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-139">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-140">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-141">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-143">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-143">Office on Windows</span></span><br><span data-ttu-id="fe20a-144">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-145">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-145">- TaskPane</span></span><br><span data-ttu-id="fe20a-146">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-146">
        - Content</span></span><br><span data-ttu-id="fe20a-147">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="fe20a-147">
        - Custom Functions</span></span><br><span data-ttu-id="fe20a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="fe20a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fe20a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fe20a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fe20a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fe20a-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fe20a-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fe20a-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fe20a-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fe20a-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fe20a-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="fe20a-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="fe20a-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="fe20a-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fe20a-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="fe20a-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-163">
        - BindingEvents</span></span><br><span data-ttu-id="fe20a-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-164">
        - CompressedFile</span></span><br><span data-ttu-id="fe20a-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-165">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-166">
        - File</span></span><br><span data-ttu-id="fe20a-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-167">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-169">
        - Selection</span></span><br><span data-ttu-id="fe20a-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-170">
        - Settings</span></span><br><span data-ttu-id="fe20a-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-171">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-172">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-173">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-175">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-175">Office 2019 on Windows</span></span><br><span data-ttu-id="fe20a-176">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fe20a-177">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-177">- TaskPane</span></span><br><span data-ttu-id="fe20a-178">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-178">
        - Content</span></span><br><span data-ttu-id="fe20a-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="fe20a-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fe20a-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fe20a-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fe20a-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fe20a-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fe20a-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fe20a-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fe20a-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fe20a-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="fe20a-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-190">- BindingEvents</span></span><br><span data-ttu-id="fe20a-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-191">
        - CompressedFile</span></span><br><span data-ttu-id="fe20a-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-192">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-193">
        - File</span></span><br><span data-ttu-id="fe20a-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-194">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-196">
        - Selection</span></span><br><span data-ttu-id="fe20a-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-197">
        - Settings</span></span><br><span data-ttu-id="fe20a-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-198">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-199">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-200">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-202">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-202">Office 2016 on Windows</span></span><br><span data-ttu-id="fe20a-203">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fe20a-204">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-204">- TaskPane</span></span><br><span data-ttu-id="fe20a-205">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-205">
        - Content</span></span></td>
    <td><span data-ttu-id="fe20a-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fe20a-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fe20a-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fe20a-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="fe20a-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-209">- BindingEvents</span></span><br><span data-ttu-id="fe20a-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-210">
        - CompressedFile</span></span><br><span data-ttu-id="fe20a-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-211">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-212">
        - File</span></span><br><span data-ttu-id="fe20a-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-213">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-215">
        - Selection</span></span><br><span data-ttu-id="fe20a-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-216">
        - Settings</span></span><br><span data-ttu-id="fe20a-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-217">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-218">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-219">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-221">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-221">Office 2013 on Windows</span></span><br><span data-ttu-id="fe20a-222">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fe20a-223">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-223">
        - TaskPane</span></span><br><span data-ttu-id="fe20a-224">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="fe20a-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fe20a-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="fe20a-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="fe20a-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-227">
        - BindingEvents</span></span><br><span data-ttu-id="fe20a-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-228">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-229">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-229">
        - File</span></span><br><span data-ttu-id="fe20a-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-230">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-232">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-232">
        - Selection</span></span><br><span data-ttu-id="fe20a-233">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-233">
        - Settings</span></span><br><span data-ttu-id="fe20a-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-234">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-235">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-236">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-238">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="fe20a-238">Office on iPad</span></span><br><span data-ttu-id="fe20a-239">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="fe20a-240">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-240">- TaskPane</span></span><br><span data-ttu-id="fe20a-241">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-241">
        - Content</span></span></td>
    <td><span data-ttu-id="fe20a-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fe20a-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fe20a-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fe20a-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fe20a-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fe20a-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fe20a-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fe20a-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fe20a-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="fe20a-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="fe20a-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="fe20a-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="fe20a-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-255">- BindingEvents</span></span><br><span data-ttu-id="fe20a-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-256">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-257">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-257">
        - File</span></span><br><span data-ttu-id="fe20a-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-258">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-260">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-260">
        - Selection</span></span><br><span data-ttu-id="fe20a-261">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-261">
        - Settings</span></span><br><span data-ttu-id="fe20a-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-262">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-263">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-264">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-266">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-266">Office on Mac</span></span><br><span data-ttu-id="fe20a-267">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="fe20a-268">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-268">- TaskPane</span></span><br><span data-ttu-id="fe20a-269">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-269">
        - Content</span></span><br><span data-ttu-id="fe20a-270">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="fe20a-270">
        - Custom Functions</span></span><br><span data-ttu-id="fe20a-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="fe20a-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fe20a-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fe20a-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fe20a-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fe20a-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fe20a-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fe20a-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fe20a-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fe20a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="fe20a-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="fe20a-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="fe20a-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fe20a-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="fe20a-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-286">- BindingEvents</span></span><br><span data-ttu-id="fe20a-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-287">
        - CompressedFile</span></span><br><span data-ttu-id="fe20a-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-288">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-289">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-289">
        - File</span></span><br><span data-ttu-id="fe20a-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-290">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-292">
        - PdfFile</span></span><br><span data-ttu-id="fe20a-293">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-293">
        - Selection</span></span><br><span data-ttu-id="fe20a-294">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-294">
        - Settings</span></span><br><span data-ttu-id="fe20a-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-295">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-296">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-297">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-299">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-299">Office 2019 on Mac</span></span><br><span data-ttu-id="fe20a-300">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fe20a-301">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-301">- TaskPane</span></span><br><span data-ttu-id="fe20a-302">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-302">
        - Content</span></span><br><span data-ttu-id="fe20a-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="fe20a-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fe20a-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fe20a-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fe20a-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fe20a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fe20a-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fe20a-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fe20a-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fe20a-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="fe20a-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-314">- BindingEvents</span></span><br><span data-ttu-id="fe20a-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-315">
        - CompressedFile</span></span><br><span data-ttu-id="fe20a-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-316">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-317">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-317">
        - File</span></span><br><span data-ttu-id="fe20a-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-318">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-320">
        - PdfFile</span></span><br><span data-ttu-id="fe20a-321">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-321">
        - Selection</span></span><br><span data-ttu-id="fe20a-322">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-322">
        - Settings</span></span><br><span data-ttu-id="fe20a-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-323">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-324">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-325">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-327">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-327">Office 2016 on Mac</span></span><br><span data-ttu-id="fe20a-328">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fe20a-329">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-329">- TaskPane</span></span><br><span data-ttu-id="fe20a-330">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-330">
        - Content</span></span></td>
    <td><span data-ttu-id="fe20a-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fe20a-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fe20a-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fe20a-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="fe20a-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-334">- BindingEvents</span></span><br><span data-ttu-id="fe20a-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-335">
        - CompressedFile</span></span><br><span data-ttu-id="fe20a-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-336">
        - DocumentEvents</span></span><br><span data-ttu-id="fe20a-337">
        - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-337">
        - File</span></span><br><span data-ttu-id="fe20a-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-338">
        - MatrixBindings</span></span><br><span data-ttu-id="fe20a-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-340">
        - PdfFile</span></span><br><span data-ttu-id="fe20a-341">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-341">
        - Selection</span></span><br><span data-ttu-id="fe20a-342">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-342">
        - Settings</span></span><br><span data-ttu-id="fe20a-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-343">
        - TableBindings</span></span><br><span data-ttu-id="fe20a-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-344">
        - TableCoercion</span></span><br><span data-ttu-id="fe20a-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-345">
        - TextBindings</span></span><br><span data-ttu-id="fe20a-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="fe20a-347">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="fe20a-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="fe20a-348">Пользовательские функции (только Excel)</span><span class="sxs-lookup"><span data-stu-id="fe20a-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="fe20a-349">Платформа</span><span class="sxs-lookup"><span data-stu-id="fe20a-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="fe20a-350">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="fe20a-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="fe20a-351">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="fe20a-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="fe20a-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="fe20a-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-353">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fe20a-353">Office on the web</span></span></td>
    <td><span data-ttu-id="fe20a-354">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="fe20a-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="fe20a-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-356">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-356">Office on Windows</span></span><br><span data-ttu-id="fe20a-357">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="fe20a-358">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="fe20a-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="fe20a-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-360">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-360">Office on Mac</span></span><br><span data-ttu-id="fe20a-361">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="fe20a-362">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="fe20a-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="fe20a-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="fe20a-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="fe20a-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fe20a-365">Платформа</span><span class="sxs-lookup"><span data-stu-id="fe20a-365">Platform</span></span></th>
    <th><span data-ttu-id="fe20a-366">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="fe20a-366">Extension points</span></span></th>
    <th><span data-ttu-id="fe20a-367">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="fe20a-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="fe20a-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="fe20a-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-369">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fe20a-369">Office on the web</span></span><br><span data-ttu-id="fe20a-370">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="fe20a-370">(modern)</span></span></td>
    <td> <span data-ttu-id="fe20a-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fe20a-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fe20a-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fe20a-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="fe20a-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="fe20a-384">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-385">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fe20a-385">Office on the web</span></span><br><span data-ttu-id="fe20a-386">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="fe20a-386">(classic)</span></span></td>
    <td> <span data-ttu-id="fe20a-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fe20a-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fe20a-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="fe20a-398">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-399">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-399">Office on Windows</span></span><br><span data-ttu-id="fe20a-400">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fe20a-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="fe20a-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="fe20a-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fe20a-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fe20a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="fe20a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="fe20a-415">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-416">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-416">Office 2019 on Windows</span></span><br><span data-ttu-id="fe20a-417">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fe20a-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="fe20a-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="fe20a-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fe20a-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fe20a-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="fe20a-431">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-432">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-432">Office 2016 on Windows</span></span><br><span data-ttu-id="fe20a-433">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fe20a-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="fe20a-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="fe20a-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="fe20a-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="fe20a-444">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-445">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-445">Office 2013 on Windows</span></span><br><span data-ttu-id="fe20a-446">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="fe20a-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="fe20a-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="fe20a-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="fe20a-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="fe20a-455">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-456">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="fe20a-456">Office on iOS</span></span><br><span data-ttu-id="fe20a-457">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="fe20a-465">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-466">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-466">Office on Mac</span></span><br><span data-ttu-id="fe20a-467">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fe20a-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fe20a-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fe20a-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="fe20a-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="fe20a-481">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-482">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-482">Office 2019 on Mac</span></span><br><span data-ttu-id="fe20a-483">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fe20a-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fe20a-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="fe20a-495">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-496">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-496">Office 2016 on Mac</span></span><br><span data-ttu-id="fe20a-497">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fe20a-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fe20a-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fe20a-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fe20a-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="fe20a-509">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-510">Office для Android</span><span class="sxs-lookup"><span data-stu-id="fe20a-510">Office on Android</span></span><br><span data-ttu-id="fe20a-511">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fe20a-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Организатор встречи (создание): собрание по сети</a> (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="fe20a-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="fe20a-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fe20a-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fe20a-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fe20a-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fe20a-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="fe20a-520">Недоступно</span><span class="sxs-lookup"><span data-stu-id="fe20a-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="fe20a-521">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="fe20a-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe20a-522">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="fe20a-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="fe20a-523">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="fe20a-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="fe20a-524">Word</span><span class="sxs-lookup"><span data-stu-id="fe20a-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fe20a-525">Платформа</span><span class="sxs-lookup"><span data-stu-id="fe20a-525">Platform</span></span></th>
    <th><span data-ttu-id="fe20a-526">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="fe20a-526">Extension points</span></span></th>
    <th><span data-ttu-id="fe20a-527">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="fe20a-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="fe20a-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="fe20a-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-529">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fe20a-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="fe20a-530">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-530">- TaskPane</span></span><br><span data-ttu-id="fe20a-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fe20a-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fe20a-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fe20a-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fe20a-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="fe20a-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-538">- BindingEvents</span></span><br><span data-ttu-id="fe20a-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-540">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-541">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-541">
         - File</span></span><br><span data-ttu-id="fe20a-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-543">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-546">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-547">
         - Selection</span></span><br><span data-ttu-id="fe20a-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-548">
         - Settings</span></span><br><span data-ttu-id="fe20a-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-549">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-550">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-551">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-552">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-554">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-554">Office on Windows</span></span><br><span data-ttu-id="fe20a-555">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-556">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-556">- TaskPane</span></span><br><span data-ttu-id="fe20a-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fe20a-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fe20a-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fe20a-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fe20a-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="fe20a-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-564">- BindingEvents</span></span><br><span data-ttu-id="fe20a-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-565">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-567">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-568">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-568">
         - File</span></span><br><span data-ttu-id="fe20a-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-570">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-573">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-574">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-574">
         - Selection</span></span><br><span data-ttu-id="fe20a-575">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-575">
         - Settings</span></span><br><span data-ttu-id="fe20a-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-576">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-577">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-578">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-579">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-581">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-581">Office 2019 on Windows</span></span><br><span data-ttu-id="fe20a-582">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-583">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-583">- TaskPane</span></span><br><span data-ttu-id="fe20a-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fe20a-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fe20a-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fe20a-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-590">- BindingEvents</span></span><br><span data-ttu-id="fe20a-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-591">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-593">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-594">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-594">
         - File</span></span><br><span data-ttu-id="fe20a-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-596">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-599">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-600">
         - Selection</span></span><br><span data-ttu-id="fe20a-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-601">
         - Settings</span></span><br><span data-ttu-id="fe20a-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-602">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-603">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-604">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-605">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-607">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-607">Office 2016 on Windows</span></span><br><span data-ttu-id="fe20a-608">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-609">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fe20a-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fe20a-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fe20a-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-613">- BindingEvents</span></span><br><span data-ttu-id="fe20a-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-614">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-616">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-617">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-617">
         - File</span></span><br><span data-ttu-id="fe20a-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-619">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-622">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-623">
         - Selection</span></span><br><span data-ttu-id="fe20a-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-624">
         - Settings</span></span><br><span data-ttu-id="fe20a-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-625">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-626">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-627">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-628">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-630">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-630">Office 2013 on Windows</span></span><br><span data-ttu-id="fe20a-631">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-632">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fe20a-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="fe20a-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-635">- BindingEvents</span></span><br><span data-ttu-id="fe20a-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-636">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-638">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-639">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-639">
         - File</span></span><br><span data-ttu-id="fe20a-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-641">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-644">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-645">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-645">
         - Selection</span></span><br><span data-ttu-id="fe20a-646">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-646">
         - Settings</span></span><br><span data-ttu-id="fe20a-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-647">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-648">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-649">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-650">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-652">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="fe20a-652">Office on iPad</span></span><br><span data-ttu-id="fe20a-653">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-654">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fe20a-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fe20a-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fe20a-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="fe20a-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-660">- BindingEvents</span></span><br><span data-ttu-id="fe20a-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-661">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-663">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-664">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-664">
         - File</span></span><br><span data-ttu-id="fe20a-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-666">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-669">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-670">
         - Selection</span></span><br><span data-ttu-id="fe20a-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-671">
         - Settings</span></span><br><span data-ttu-id="fe20a-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-672">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-673">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-674">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-675">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-677">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-677">Office on Mac</span></span><br><span data-ttu-id="fe20a-678">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-679">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-679">- TaskPane</span></span><br><span data-ttu-id="fe20a-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fe20a-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fe20a-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fe20a-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fe20a-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="fe20a-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-687">- BindingEvents</span></span><br><span data-ttu-id="fe20a-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-688">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-690">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-691">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-691">
         - File</span></span><br><span data-ttu-id="fe20a-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-693">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-696">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-697">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-697">
         - Selection</span></span><br><span data-ttu-id="fe20a-698">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-698">
         - Settings</span></span><br><span data-ttu-id="fe20a-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-699">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-700">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-701">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-702">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-704">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-704">Office 2019 on Mac</span></span><br><span data-ttu-id="fe20a-705">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-706">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-706">- TaskPane</span></span><br><span data-ttu-id="fe20a-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fe20a-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fe20a-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fe20a-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="fe20a-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-713">- BindingEvents</span></span><br><span data-ttu-id="fe20a-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-714">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-716">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-717">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-717">
         - File</span></span><br><span data-ttu-id="fe20a-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-719">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-722">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-723">
         - Selection</span></span><br><span data-ttu-id="fe20a-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-724">
         - Settings</span></span><br><span data-ttu-id="fe20a-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-725">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-726">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-727">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-728">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-730">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-730">Office 2016 on Mac</span></span><br><span data-ttu-id="fe20a-731">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-732">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fe20a-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fe20a-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fe20a-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-736">- BindingEvents</span></span><br><span data-ttu-id="fe20a-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-737">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fe20a-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="fe20a-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-739">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-740">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="fe20a-740">
         - File</span></span><br><span data-ttu-id="fe20a-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-742">
         - MatrixBindings</span></span><br><span data-ttu-id="fe20a-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="fe20a-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fe20a-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-745">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-746">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-746">
         - Selection</span></span><br><span data-ttu-id="fe20a-747">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fe20a-747">
         - Settings</span></span><br><span data-ttu-id="fe20a-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-748">
         - TableBindings</span></span><br><span data-ttu-id="fe20a-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-749">
         - TableCoercion</span></span><br><span data-ttu-id="fe20a-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fe20a-750">
         - TextBindings</span></span><br><span data-ttu-id="fe20a-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-751">
         - TextCoercion</span></span><br><span data-ttu-id="fe20a-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="fe20a-753">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="fe20a-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="fe20a-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="fe20a-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fe20a-755">Платформа</span><span class="sxs-lookup"><span data-stu-id="fe20a-755">Platform</span></span></th>
    <th><span data-ttu-id="fe20a-756">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="fe20a-756">Extension points</span></span></th>
    <th><span data-ttu-id="fe20a-757">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="fe20a-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="fe20a-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="fe20a-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-759">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fe20a-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="fe20a-760">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-760">- Content</span></span><br><span data-ttu-id="fe20a-761">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-761">
         - TaskPane</span></span><br><span data-ttu-id="fe20a-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="fe20a-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fe20a-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="fe20a-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-767">- ActiveView</span></span><br><span data-ttu-id="fe20a-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-768">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-769">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-770">
         - File</span></span><br><span data-ttu-id="fe20a-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-771">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-772">
         - Selection</span></span><br><span data-ttu-id="fe20a-773">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-773">
         - Settings</span></span><br><span data-ttu-id="fe20a-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-775">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-775">Office on Windows</span></span><br><span data-ttu-id="fe20a-776">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-777">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-777">- Content</span></span><br><span data-ttu-id="fe20a-778">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-778">
         - TaskPane</span></span><br><span data-ttu-id="fe20a-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="fe20a-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fe20a-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="fe20a-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-784">- ActiveView</span></span><br><span data-ttu-id="fe20a-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-785">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-786">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-787">
         - File</span></span><br><span data-ttu-id="fe20a-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-788">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-789">
         - Selection</span></span><br><span data-ttu-id="fe20a-790">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-790">
         - Settings</span></span><br><span data-ttu-id="fe20a-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-792">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-792">Office 2019 on Windows</span></span><br><span data-ttu-id="fe20a-793">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-794">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-794">- Content</span></span><br><span data-ttu-id="fe20a-795">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-795">
         - TaskPane</span></span><br><span data-ttu-id="fe20a-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-799">- ActiveView</span></span><br><span data-ttu-id="fe20a-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-800">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-801">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-802">
         - File</span></span><br><span data-ttu-id="fe20a-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-803">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-804">
         - Selection</span></span><br><span data-ttu-id="fe20a-805">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-805">
         - Settings</span></span><br><span data-ttu-id="fe20a-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-807">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-807">Office 2016 on Windows</span></span><br><span data-ttu-id="fe20a-808">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-809">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-809">- Content</span></span><br><span data-ttu-id="fe20a-810">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fe20a-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="fe20a-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-813">- ActiveView</span></span><br><span data-ttu-id="fe20a-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-814">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-815">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-816">
         - File</span></span><br><span data-ttu-id="fe20a-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-817">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-818">
         - Selection</span></span><br><span data-ttu-id="fe20a-819">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-819">
         - Settings</span></span><br><span data-ttu-id="fe20a-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-821">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-821">Office 2013 on Windows</span></span><br><span data-ttu-id="fe20a-822">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-823">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-823">- Content</span></span><br><span data-ttu-id="fe20a-824">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="fe20a-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fe20a-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="fe20a-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-827">- ActiveView</span></span><br><span data-ttu-id="fe20a-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-828">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-829">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-830">
         - File</span></span><br><span data-ttu-id="fe20a-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-831">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-832">
         - Selection</span></span><br><span data-ttu-id="fe20a-833">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-833">
         - Settings</span></span><br><span data-ttu-id="fe20a-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-835">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="fe20a-835">Office on iPad</span></span><br><span data-ttu-id="fe20a-836">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-837">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-837">- Content</span></span><br><span data-ttu-id="fe20a-838">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="fe20a-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-842">- ActiveView</span></span><br><span data-ttu-id="fe20a-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-843">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-844">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-845">
         - File</span></span><br><span data-ttu-id="fe20a-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-846">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-847">
         - Selection</span></span><br><span data-ttu-id="fe20a-848">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-848">
         - Settings</span></span><br><span data-ttu-id="fe20a-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-850">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-850">Office on Mac</span></span><br><span data-ttu-id="fe20a-851">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="fe20a-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="fe20a-852">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-852">- Content</span></span><br><span data-ttu-id="fe20a-853">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-853">
         - TaskPane</span></span><br><span data-ttu-id="fe20a-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="fe20a-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fe20a-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="fe20a-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-859">- ActiveView</span></span><br><span data-ttu-id="fe20a-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-860">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-861">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-862">
         - File</span></span><br><span data-ttu-id="fe20a-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-863">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-864">
         - Selection</span></span><br><span data-ttu-id="fe20a-865">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-865">
         - Settings</span></span><br><span data-ttu-id="fe20a-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-867">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-867">Office 2019 on Mac</span></span><br><span data-ttu-id="fe20a-868">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-869">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-869">- Content</span></span><br><span data-ttu-id="fe20a-870">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-870">
         - TaskPane</span></span><br><span data-ttu-id="fe20a-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-874">- ActiveView</span></span><br><span data-ttu-id="fe20a-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-875">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-876">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-877">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-877">
         - File</span></span><br><span data-ttu-id="fe20a-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-878">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-879">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-879">
         - Selection</span></span><br><span data-ttu-id="fe20a-880">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-880">
         - Settings</span></span><br><span data-ttu-id="fe20a-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-882">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-882">Office 2016 on Mac</span></span><br><span data-ttu-id="fe20a-883">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-884">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-884">- Content</span></span><br><span data-ttu-id="fe20a-885">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fe20a-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="fe20a-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fe20a-888">- ActiveView</span></span><br><span data-ttu-id="fe20a-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-889">
         - CompressedFile</span></span><br><span data-ttu-id="fe20a-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-890">
         - DocumentEvents</span></span><br><span data-ttu-id="fe20a-891">
         - File</span><span class="sxs-lookup"><span data-stu-id="fe20a-891">
         - File</span></span><br><span data-ttu-id="fe20a-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fe20a-892">
         - PdfFile</span></span><br><span data-ttu-id="fe20a-893">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-893">
         - Selection</span></span><br><span data-ttu-id="fe20a-894">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-894">
         - Settings</span></span><br><span data-ttu-id="fe20a-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="fe20a-896">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="fe20a-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="fe20a-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="fe20a-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fe20a-898">Платформа</span><span class="sxs-lookup"><span data-stu-id="fe20a-898">Platform</span></span></th>
    <th><span data-ttu-id="fe20a-899">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="fe20a-899">Extension points</span></span></th>
    <th><span data-ttu-id="fe20a-900">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="fe20a-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="fe20a-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="fe20a-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-902">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fe20a-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="fe20a-903">- Контент</span><span class="sxs-lookup"><span data-stu-id="fe20a-903">- Content</span></span><br><span data-ttu-id="fe20a-904">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-904">
         - TaskPane</span></span><br><span data-ttu-id="fe20a-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fe20a-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="fe20a-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fe20a-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fe20a-909">- DocumentEvents</span></span><br><span data-ttu-id="fe20a-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="fe20a-911">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="fe20a-911">
         - Settings</span></span><br><span data-ttu-id="fe20a-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="fe20a-913">Project</span><span class="sxs-lookup"><span data-stu-id="fe20a-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fe20a-914">Платформа</span><span class="sxs-lookup"><span data-stu-id="fe20a-914">Platform</span></span></th>
    <th><span data-ttu-id="fe20a-915">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="fe20a-915">Extension points</span></span></th>
    <th><span data-ttu-id="fe20a-916">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="fe20a-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="fe20a-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="fe20a-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-918">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-918">Office 2019 on Windows</span></span><br><span data-ttu-id="fe20a-919">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-920">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-922">- Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-922">- Selection</span></span><br><span data-ttu-id="fe20a-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-924">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-924">Office 2016 on Windows</span></span><br><span data-ttu-id="fe20a-925">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-926">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-928">- Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-928">- Selection</span></span><br><span data-ttu-id="fe20a-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fe20a-930">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="fe20a-930">Office 2013 on Windows</span></span><br><span data-ttu-id="fe20a-931">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="fe20a-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="fe20a-932">- Область задач</span><span class="sxs-lookup"><span data-stu-id="fe20a-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fe20a-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fe20a-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fe20a-934">- Selection</span><span class="sxs-lookup"><span data-stu-id="fe20a-934">- Selection</span></span><br><span data-ttu-id="fe20a-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fe20a-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="fe20a-936">См. также</span><span class="sxs-lookup"><span data-stu-id="fe20a-936">See also</span></span>

- [<span data-ttu-id="fe20a-937">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="fe20a-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="fe20a-938">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="fe20a-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="fe20a-939">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="fe20a-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="fe20a-940">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="fe20a-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="fe20a-941">Справочная документация по API</span><span class="sxs-lookup"><span data-stu-id="fe20a-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="fe20a-942">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="fe20a-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="fe20a-943">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="fe20a-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="fe20a-944">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="fe20a-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="fe20a-945">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="fe20a-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="fe20a-946">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="fe20a-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="fe20a-947">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="fe20a-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="fe20a-948">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="fe20a-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)