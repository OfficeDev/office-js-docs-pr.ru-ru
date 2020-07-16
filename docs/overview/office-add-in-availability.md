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
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="de7d0-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="de7d0-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="de7d0-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="de7d0-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="de7d0-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="de7d0-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="de7d0-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="de7d0-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="de7d0-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="de7d0-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="de7d0-108">Excel</span><span class="sxs-lookup"><span data-stu-id="de7d0-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="de7d0-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="de7d0-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="de7d0-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="de7d0-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="de7d0-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="de7d0-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="de7d0-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="de7d0-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="de7d0-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="de7d0-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-114">- TaskPane</span></span><br><span data-ttu-id="de7d0-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-115">
        - Content</span></span><br><span data-ttu-id="de7d0-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="de7d0-116">
        - Custom Functions</span></span><br><span data-ttu-id="de7d0-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="de7d0-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="de7d0-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="de7d0-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="de7d0-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="de7d0-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="de7d0-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="de7d0-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="de7d0-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="de7d0-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="de7d0-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="de7d0-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="de7d0-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="de7d0-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="de7d0-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="de7d0-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-131">
        - BindingEvents</span></span><br><span data-ttu-id="de7d0-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-132">
        - CompressedFile</span></span><br><span data-ttu-id="de7d0-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-133">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-134">
        - File</span></span><br><span data-ttu-id="de7d0-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-135">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-137">
        - Selection</span></span><br><span data-ttu-id="de7d0-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-138">
        - Settings</span></span><br><span data-ttu-id="de7d0-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-139">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-140">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-141">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-143">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-143">Office on Windows</span></span><br><span data-ttu-id="de7d0-144">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-145">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-145">- TaskPane</span></span><br><span data-ttu-id="de7d0-146">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-146">
        - Content</span></span><br><span data-ttu-id="de7d0-147">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="de7d0-147">
        - Custom Functions</span></span><br><span data-ttu-id="de7d0-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="de7d0-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="de7d0-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="de7d0-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="de7d0-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="de7d0-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="de7d0-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="de7d0-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="de7d0-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="de7d0-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="de7d0-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="de7d0-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="de7d0-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="de7d0-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="de7d0-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="de7d0-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-163">
        - BindingEvents</span></span><br><span data-ttu-id="de7d0-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-164">
        - CompressedFile</span></span><br><span data-ttu-id="de7d0-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-165">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-166">
        - File</span></span><br><span data-ttu-id="de7d0-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-167">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-169">
        - Selection</span></span><br><span data-ttu-id="de7d0-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-170">
        - Settings</span></span><br><span data-ttu-id="de7d0-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-171">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-172">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-173">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-175">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-175">Office 2019 on Windows</span></span><br><span data-ttu-id="de7d0-176">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="de7d0-177">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-177">- TaskPane</span></span><br><span data-ttu-id="de7d0-178">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-178">
        - Content</span></span><br><span data-ttu-id="de7d0-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="de7d0-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="de7d0-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="de7d0-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="de7d0-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="de7d0-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="de7d0-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="de7d0-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="de7d0-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="de7d0-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="de7d0-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-190">- BindingEvents</span></span><br><span data-ttu-id="de7d0-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-191">
        - CompressedFile</span></span><br><span data-ttu-id="de7d0-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-192">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-193">
        - File</span></span><br><span data-ttu-id="de7d0-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-194">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-196">
        - Selection</span></span><br><span data-ttu-id="de7d0-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-197">
        - Settings</span></span><br><span data-ttu-id="de7d0-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-198">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-199">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-200">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-202">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-202">Office 2016 on Windows</span></span><br><span data-ttu-id="de7d0-203">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="de7d0-204">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-204">- TaskPane</span></span><br><span data-ttu-id="de7d0-205">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-205">
        - Content</span></span></td>
    <td><span data-ttu-id="de7d0-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="de7d0-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="de7d0-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="de7d0-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="de7d0-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-209">- BindingEvents</span></span><br><span data-ttu-id="de7d0-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-210">
        - CompressedFile</span></span><br><span data-ttu-id="de7d0-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-211">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-212">
        - File</span></span><br><span data-ttu-id="de7d0-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-213">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-215">
        - Selection</span></span><br><span data-ttu-id="de7d0-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-216">
        - Settings</span></span><br><span data-ttu-id="de7d0-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-217">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-218">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-219">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-221">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-221">Office 2013 on Windows</span></span><br><span data-ttu-id="de7d0-222">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="de7d0-223">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-223">
        - TaskPane</span></span><br><span data-ttu-id="de7d0-224">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="de7d0-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="de7d0-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="de7d0-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="de7d0-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-227">
        - BindingEvents</span></span><br><span data-ttu-id="de7d0-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-228">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-229">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-229">
        - File</span></span><br><span data-ttu-id="de7d0-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-230">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-232">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-232">
        - Selection</span></span><br><span data-ttu-id="de7d0-233">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-233">
        - Settings</span></span><br><span data-ttu-id="de7d0-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-234">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-235">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-236">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-238">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="de7d0-238">Office on iPad</span></span><br><span data-ttu-id="de7d0-239">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="de7d0-240">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-240">- TaskPane</span></span><br><span data-ttu-id="de7d0-241">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-241">
        - Content</span></span></td>
    <td><span data-ttu-id="de7d0-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="de7d0-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="de7d0-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="de7d0-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="de7d0-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="de7d0-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="de7d0-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="de7d0-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="de7d0-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="de7d0-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="de7d0-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="de7d0-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="de7d0-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-255">- BindingEvents</span></span><br><span data-ttu-id="de7d0-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-256">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-257">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-257">
        - File</span></span><br><span data-ttu-id="de7d0-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-258">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-260">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-260">
        - Selection</span></span><br><span data-ttu-id="de7d0-261">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-261">
        - Settings</span></span><br><span data-ttu-id="de7d0-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-262">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-263">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-264">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-266">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-266">Office on Mac</span></span><br><span data-ttu-id="de7d0-267">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="de7d0-268">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-268">- TaskPane</span></span><br><span data-ttu-id="de7d0-269">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-269">
        - Content</span></span><br><span data-ttu-id="de7d0-270">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="de7d0-270">
        - Custom Functions</span></span><br><span data-ttu-id="de7d0-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="de7d0-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="de7d0-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="de7d0-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="de7d0-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="de7d0-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="de7d0-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="de7d0-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="de7d0-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="de7d0-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="de7d0-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="de7d0-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="de7d0-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="de7d0-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="de7d0-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-286">- BindingEvents</span></span><br><span data-ttu-id="de7d0-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-287">
        - CompressedFile</span></span><br><span data-ttu-id="de7d0-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-288">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-289">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-289">
        - File</span></span><br><span data-ttu-id="de7d0-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-290">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-292">
        - PdfFile</span></span><br><span data-ttu-id="de7d0-293">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-293">
        - Selection</span></span><br><span data-ttu-id="de7d0-294">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-294">
        - Settings</span></span><br><span data-ttu-id="de7d0-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-295">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-296">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-297">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-299">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-299">Office 2019 on Mac</span></span><br><span data-ttu-id="de7d0-300">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="de7d0-301">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-301">- TaskPane</span></span><br><span data-ttu-id="de7d0-302">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-302">
        - Content</span></span><br><span data-ttu-id="de7d0-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="de7d0-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="de7d0-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="de7d0-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="de7d0-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="de7d0-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="de7d0-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="de7d0-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="de7d0-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="de7d0-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="de7d0-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-314">- BindingEvents</span></span><br><span data-ttu-id="de7d0-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-315">
        - CompressedFile</span></span><br><span data-ttu-id="de7d0-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-316">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-317">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-317">
        - File</span></span><br><span data-ttu-id="de7d0-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-318">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-320">
        - PdfFile</span></span><br><span data-ttu-id="de7d0-321">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-321">
        - Selection</span></span><br><span data-ttu-id="de7d0-322">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-322">
        - Settings</span></span><br><span data-ttu-id="de7d0-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-323">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-324">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-325">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-327">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-327">Office 2016 on Mac</span></span><br><span data-ttu-id="de7d0-328">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="de7d0-329">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-329">- TaskPane</span></span><br><span data-ttu-id="de7d0-330">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-330">
        - Content</span></span></td>
    <td><span data-ttu-id="de7d0-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="de7d0-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="de7d0-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="de7d0-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="de7d0-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-334">- BindingEvents</span></span><br><span data-ttu-id="de7d0-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-335">
        - CompressedFile</span></span><br><span data-ttu-id="de7d0-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-336">
        - DocumentEvents</span></span><br><span data-ttu-id="de7d0-337">
        - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-337">
        - File</span></span><br><span data-ttu-id="de7d0-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-338">
        - MatrixBindings</span></span><br><span data-ttu-id="de7d0-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-340">
        - PdfFile</span></span><br><span data-ttu-id="de7d0-341">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-341">
        - Selection</span></span><br><span data-ttu-id="de7d0-342">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-342">
        - Settings</span></span><br><span data-ttu-id="de7d0-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-343">
        - TableBindings</span></span><br><span data-ttu-id="de7d0-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-344">
        - TableCoercion</span></span><br><span data-ttu-id="de7d0-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-345">
        - TextBindings</span></span><br><span data-ttu-id="de7d0-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="de7d0-347">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="de7d0-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="de7d0-348">Пользовательские функции (только Excel)</span><span class="sxs-lookup"><span data-stu-id="de7d0-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="de7d0-349">Платформа</span><span class="sxs-lookup"><span data-stu-id="de7d0-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="de7d0-350">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="de7d0-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="de7d0-351">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="de7d0-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="de7d0-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="de7d0-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-353">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="de7d0-353">Office on the web</span></span></td>
    <td><span data-ttu-id="de7d0-354">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="de7d0-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="de7d0-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-356">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-356">Office on Windows</span></span><br><span data-ttu-id="de7d0-357">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="de7d0-358">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="de7d0-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="de7d0-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-360">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-360">Office on Mac</span></span><br><span data-ttu-id="de7d0-361">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="de7d0-362">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="de7d0-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="de7d0-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="de7d0-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="de7d0-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="de7d0-365">Платформа</span><span class="sxs-lookup"><span data-stu-id="de7d0-365">Platform</span></span></th>
    <th><span data-ttu-id="de7d0-366">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="de7d0-366">Extension points</span></span></th>
    <th><span data-ttu-id="de7d0-367">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="de7d0-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="de7d0-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="de7d0-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-369">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="de7d0-369">Office on the web</span></span><br><span data-ttu-id="de7d0-370">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="de7d0-370">(modern)</span></span></td>
    <td> <span data-ttu-id="de7d0-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="de7d0-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="de7d0-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="de7d0-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="de7d0-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="de7d0-384">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-385">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="de7d0-385">Office on the web</span></span><br><span data-ttu-id="de7d0-386">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="de7d0-386">(classic)</span></span></td>
    <td> <span data-ttu-id="de7d0-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="de7d0-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="de7d0-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="de7d0-398">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-399">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-399">Office on Windows</span></span><br><span data-ttu-id="de7d0-400">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="de7d0-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="de7d0-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="de7d0-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="de7d0-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="de7d0-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="de7d0-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="de7d0-415">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-416">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-416">Office 2019 on Windows</span></span><br><span data-ttu-id="de7d0-417">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="de7d0-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="de7d0-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="de7d0-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="de7d0-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="de7d0-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="de7d0-431">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-432">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-432">Office 2016 on Windows</span></span><br><span data-ttu-id="de7d0-433">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="de7d0-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="de7d0-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="de7d0-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="de7d0-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="de7d0-444">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-445">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-445">Office 2013 on Windows</span></span><br><span data-ttu-id="de7d0-446">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="de7d0-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="de7d0-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="de7d0-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="de7d0-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="de7d0-455">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-456">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="de7d0-456">Office on iOS</span></span><br><span data-ttu-id="de7d0-457">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="de7d0-465">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-466">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-466">Office on Mac</span></span><br><span data-ttu-id="de7d0-467">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="de7d0-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="de7d0-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="de7d0-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="de7d0-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="de7d0-481">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-482">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-482">Office 2019 on Mac</span></span><br><span data-ttu-id="de7d0-483">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="de7d0-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="de7d0-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="de7d0-495">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-496">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-496">Office 2016 on Mac</span></span><br><span data-ttu-id="de7d0-497">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="de7d0-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="de7d0-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="de7d0-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="de7d0-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="de7d0-509">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-510">Office для Android</span><span class="sxs-lookup"><span data-stu-id="de7d0-510">Office on Android</span></span><br><span data-ttu-id="de7d0-511">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="de7d0-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Организатор встречи (создание): собрание по сети</a> (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="de7d0-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="de7d0-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="de7d0-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="de7d0-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="de7d0-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="de7d0-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="de7d0-520">Недоступно</span><span class="sxs-lookup"><span data-stu-id="de7d0-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="de7d0-521">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="de7d0-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="de7d0-522">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="de7d0-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="de7d0-523">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="de7d0-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="de7d0-524">Word</span><span class="sxs-lookup"><span data-stu-id="de7d0-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="de7d0-525">Платформа</span><span class="sxs-lookup"><span data-stu-id="de7d0-525">Platform</span></span></th>
    <th><span data-ttu-id="de7d0-526">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="de7d0-526">Extension points</span></span></th>
    <th><span data-ttu-id="de7d0-527">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="de7d0-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="de7d0-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="de7d0-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-529">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="de7d0-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="de7d0-530">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-530">- TaskPane</span></span><br><span data-ttu-id="de7d0-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="de7d0-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="de7d0-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="de7d0-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="de7d0-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="de7d0-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-538">- BindingEvents</span></span><br><span data-ttu-id="de7d0-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-540">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-541">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-541">
         - File</span></span><br><span data-ttu-id="de7d0-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-543">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-546">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-547">
         - Selection</span></span><br><span data-ttu-id="de7d0-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-548">
         - Settings</span></span><br><span data-ttu-id="de7d0-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-549">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-550">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-551">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-552">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-554">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-554">Office on Windows</span></span><br><span data-ttu-id="de7d0-555">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-556">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-556">- TaskPane</span></span><br><span data-ttu-id="de7d0-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="de7d0-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="de7d0-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="de7d0-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="de7d0-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="de7d0-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-564">- BindingEvents</span></span><br><span data-ttu-id="de7d0-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-565">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-567">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-568">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-568">
         - File</span></span><br><span data-ttu-id="de7d0-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-570">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-573">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-574">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-574">
         - Selection</span></span><br><span data-ttu-id="de7d0-575">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-575">
         - Settings</span></span><br><span data-ttu-id="de7d0-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-576">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-577">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-578">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-579">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-581">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-581">Office 2019 on Windows</span></span><br><span data-ttu-id="de7d0-582">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-583">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-583">- TaskPane</span></span><br><span data-ttu-id="de7d0-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="de7d0-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="de7d0-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="de7d0-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-590">- BindingEvents</span></span><br><span data-ttu-id="de7d0-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-591">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-593">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-594">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-594">
         - File</span></span><br><span data-ttu-id="de7d0-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-596">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-599">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-600">
         - Selection</span></span><br><span data-ttu-id="de7d0-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-601">
         - Settings</span></span><br><span data-ttu-id="de7d0-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-602">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-603">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-604">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-605">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-607">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-607">Office 2016 on Windows</span></span><br><span data-ttu-id="de7d0-608">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-609">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="de7d0-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="de7d0-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="de7d0-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-613">- BindingEvents</span></span><br><span data-ttu-id="de7d0-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-614">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-616">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-617">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-617">
         - File</span></span><br><span data-ttu-id="de7d0-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-619">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-622">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-623">
         - Selection</span></span><br><span data-ttu-id="de7d0-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-624">
         - Settings</span></span><br><span data-ttu-id="de7d0-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-625">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-626">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-627">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-628">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-630">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-630">Office 2013 on Windows</span></span><br><span data-ttu-id="de7d0-631">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-632">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="de7d0-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="de7d0-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-635">- BindingEvents</span></span><br><span data-ttu-id="de7d0-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-636">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-638">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-639">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-639">
         - File</span></span><br><span data-ttu-id="de7d0-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-641">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-644">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-645">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-645">
         - Selection</span></span><br><span data-ttu-id="de7d0-646">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-646">
         - Settings</span></span><br><span data-ttu-id="de7d0-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-647">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-648">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-649">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-650">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-652">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="de7d0-652">Office on iPad</span></span><br><span data-ttu-id="de7d0-653">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-654">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="de7d0-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="de7d0-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="de7d0-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="de7d0-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-660">- BindingEvents</span></span><br><span data-ttu-id="de7d0-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-661">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-663">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-664">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-664">
         - File</span></span><br><span data-ttu-id="de7d0-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-666">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-669">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-670">
         - Selection</span></span><br><span data-ttu-id="de7d0-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-671">
         - Settings</span></span><br><span data-ttu-id="de7d0-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-672">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-673">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-674">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-675">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-677">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-677">Office on Mac</span></span><br><span data-ttu-id="de7d0-678">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-679">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-679">- TaskPane</span></span><br><span data-ttu-id="de7d0-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="de7d0-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="de7d0-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="de7d0-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="de7d0-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="de7d0-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-687">- BindingEvents</span></span><br><span data-ttu-id="de7d0-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-688">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-690">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-691">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-691">
         - File</span></span><br><span data-ttu-id="de7d0-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-693">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-696">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-697">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-697">
         - Selection</span></span><br><span data-ttu-id="de7d0-698">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-698">
         - Settings</span></span><br><span data-ttu-id="de7d0-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-699">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-700">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-701">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-702">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-704">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-704">Office 2019 on Mac</span></span><br><span data-ttu-id="de7d0-705">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-706">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-706">- TaskPane</span></span><br><span data-ttu-id="de7d0-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="de7d0-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="de7d0-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="de7d0-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="de7d0-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-713">- BindingEvents</span></span><br><span data-ttu-id="de7d0-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-714">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-716">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-717">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-717">
         - File</span></span><br><span data-ttu-id="de7d0-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-719">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-722">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-723">
         - Selection</span></span><br><span data-ttu-id="de7d0-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-724">
         - Settings</span></span><br><span data-ttu-id="de7d0-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-725">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-726">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-727">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-728">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-730">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-730">Office 2016 on Mac</span></span><br><span data-ttu-id="de7d0-731">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-732">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="de7d0-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="de7d0-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="de7d0-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-736">- BindingEvents</span></span><br><span data-ttu-id="de7d0-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-737">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="de7d0-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="de7d0-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-739">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-740">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="de7d0-740">
         - File</span></span><br><span data-ttu-id="de7d0-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-742">
         - MatrixBindings</span></span><br><span data-ttu-id="de7d0-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="de7d0-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="de7d0-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-745">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-746">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-746">
         - Selection</span></span><br><span data-ttu-id="de7d0-747">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="de7d0-747">
         - Settings</span></span><br><span data-ttu-id="de7d0-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-748">
         - TableBindings</span></span><br><span data-ttu-id="de7d0-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-749">
         - TableCoercion</span></span><br><span data-ttu-id="de7d0-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="de7d0-750">
         - TextBindings</span></span><br><span data-ttu-id="de7d0-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-751">
         - TextCoercion</span></span><br><span data-ttu-id="de7d0-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="de7d0-753">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="de7d0-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="de7d0-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="de7d0-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="de7d0-755">Платформа</span><span class="sxs-lookup"><span data-stu-id="de7d0-755">Platform</span></span></th>
    <th><span data-ttu-id="de7d0-756">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="de7d0-756">Extension points</span></span></th>
    <th><span data-ttu-id="de7d0-757">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="de7d0-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="de7d0-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="de7d0-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-759">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="de7d0-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="de7d0-760">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-760">- Content</span></span><br><span data-ttu-id="de7d0-761">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-761">
         - TaskPane</span></span><br><span data-ttu-id="de7d0-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="de7d0-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="de7d0-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="de7d0-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-767">- ActiveView</span></span><br><span data-ttu-id="de7d0-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-768">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-769">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-770">
         - File</span></span><br><span data-ttu-id="de7d0-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-771">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-772">
         - Selection</span></span><br><span data-ttu-id="de7d0-773">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-773">
         - Settings</span></span><br><span data-ttu-id="de7d0-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-775">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-775">Office on Windows</span></span><br><span data-ttu-id="de7d0-776">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-777">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-777">- Content</span></span><br><span data-ttu-id="de7d0-778">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-778">
         - TaskPane</span></span><br><span data-ttu-id="de7d0-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="de7d0-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="de7d0-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="de7d0-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-784">- ActiveView</span></span><br><span data-ttu-id="de7d0-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-785">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-786">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-787">
         - File</span></span><br><span data-ttu-id="de7d0-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-788">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-789">
         - Selection</span></span><br><span data-ttu-id="de7d0-790">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-790">
         - Settings</span></span><br><span data-ttu-id="de7d0-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-792">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-792">Office 2019 on Windows</span></span><br><span data-ttu-id="de7d0-793">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-794">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-794">- Content</span></span><br><span data-ttu-id="de7d0-795">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-795">
         - TaskPane</span></span><br><span data-ttu-id="de7d0-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-799">- ActiveView</span></span><br><span data-ttu-id="de7d0-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-800">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-801">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-802">
         - File</span></span><br><span data-ttu-id="de7d0-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-803">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-804">
         - Selection</span></span><br><span data-ttu-id="de7d0-805">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-805">
         - Settings</span></span><br><span data-ttu-id="de7d0-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-807">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-807">Office 2016 on Windows</span></span><br><span data-ttu-id="de7d0-808">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-809">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-809">- Content</span></span><br><span data-ttu-id="de7d0-810">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="de7d0-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="de7d0-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-813">- ActiveView</span></span><br><span data-ttu-id="de7d0-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-814">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-815">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-816">
         - File</span></span><br><span data-ttu-id="de7d0-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-817">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-818">
         - Selection</span></span><br><span data-ttu-id="de7d0-819">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-819">
         - Settings</span></span><br><span data-ttu-id="de7d0-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-821">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-821">Office 2013 on Windows</span></span><br><span data-ttu-id="de7d0-822">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-823">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-823">- Content</span></span><br><span data-ttu-id="de7d0-824">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="de7d0-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="de7d0-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="de7d0-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-827">- ActiveView</span></span><br><span data-ttu-id="de7d0-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-828">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-829">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-830">
         - File</span></span><br><span data-ttu-id="de7d0-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-831">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-832">
         - Selection</span></span><br><span data-ttu-id="de7d0-833">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-833">
         - Settings</span></span><br><span data-ttu-id="de7d0-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-835">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="de7d0-835">Office on iPad</span></span><br><span data-ttu-id="de7d0-836">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-837">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-837">- Content</span></span><br><span data-ttu-id="de7d0-838">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="de7d0-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-842">- ActiveView</span></span><br><span data-ttu-id="de7d0-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-843">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-844">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-845">
         - File</span></span><br><span data-ttu-id="de7d0-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-846">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-847">
         - Selection</span></span><br><span data-ttu-id="de7d0-848">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-848">
         - Settings</span></span><br><span data-ttu-id="de7d0-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-850">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-850">Office on Mac</span></span><br><span data-ttu-id="de7d0-851">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="de7d0-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="de7d0-852">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-852">- Content</span></span><br><span data-ttu-id="de7d0-853">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-853">
         - TaskPane</span></span><br><span data-ttu-id="de7d0-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="de7d0-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="de7d0-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="de7d0-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-859">- ActiveView</span></span><br><span data-ttu-id="de7d0-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-860">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-861">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-862">
         - File</span></span><br><span data-ttu-id="de7d0-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-863">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-864">
         - Selection</span></span><br><span data-ttu-id="de7d0-865">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-865">
         - Settings</span></span><br><span data-ttu-id="de7d0-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-867">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-867">Office 2019 on Mac</span></span><br><span data-ttu-id="de7d0-868">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-869">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-869">- Content</span></span><br><span data-ttu-id="de7d0-870">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-870">
         - TaskPane</span></span><br><span data-ttu-id="de7d0-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-874">- ActiveView</span></span><br><span data-ttu-id="de7d0-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-875">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-876">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-877">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-877">
         - File</span></span><br><span data-ttu-id="de7d0-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-878">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-879">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-879">
         - Selection</span></span><br><span data-ttu-id="de7d0-880">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-880">
         - Settings</span></span><br><span data-ttu-id="de7d0-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-882">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-882">Office 2016 on Mac</span></span><br><span data-ttu-id="de7d0-883">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-884">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-884">- Content</span></span><br><span data-ttu-id="de7d0-885">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="de7d0-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="de7d0-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="de7d0-888">- ActiveView</span></span><br><span data-ttu-id="de7d0-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-889">
         - CompressedFile</span></span><br><span data-ttu-id="de7d0-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-890">
         - DocumentEvents</span></span><br><span data-ttu-id="de7d0-891">
         - File</span><span class="sxs-lookup"><span data-stu-id="de7d0-891">
         - File</span></span><br><span data-ttu-id="de7d0-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="de7d0-892">
         - PdfFile</span></span><br><span data-ttu-id="de7d0-893">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-893">
         - Selection</span></span><br><span data-ttu-id="de7d0-894">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-894">
         - Settings</span></span><br><span data-ttu-id="de7d0-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="de7d0-896">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="de7d0-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="de7d0-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="de7d0-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="de7d0-898">Платформа</span><span class="sxs-lookup"><span data-stu-id="de7d0-898">Platform</span></span></th>
    <th><span data-ttu-id="de7d0-899">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="de7d0-899">Extension points</span></span></th>
    <th><span data-ttu-id="de7d0-900">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="de7d0-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="de7d0-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="de7d0-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-902">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="de7d0-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="de7d0-903">- Контент</span><span class="sxs-lookup"><span data-stu-id="de7d0-903">- Content</span></span><br><span data-ttu-id="de7d0-904">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-904">
         - TaskPane</span></span><br><span data-ttu-id="de7d0-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="de7d0-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="de7d0-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="de7d0-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="de7d0-909">- DocumentEvents</span></span><br><span data-ttu-id="de7d0-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="de7d0-911">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="de7d0-911">
         - Settings</span></span><br><span data-ttu-id="de7d0-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="de7d0-913">Project</span><span class="sxs-lookup"><span data-stu-id="de7d0-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="de7d0-914">Платформа</span><span class="sxs-lookup"><span data-stu-id="de7d0-914">Platform</span></span></th>
    <th><span data-ttu-id="de7d0-915">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="de7d0-915">Extension points</span></span></th>
    <th><span data-ttu-id="de7d0-916">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="de7d0-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="de7d0-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="de7d0-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-918">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-918">Office 2019 on Windows</span></span><br><span data-ttu-id="de7d0-919">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-920">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-922">- Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-922">- Selection</span></span><br><span data-ttu-id="de7d0-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-924">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-924">Office 2016 on Windows</span></span><br><span data-ttu-id="de7d0-925">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-926">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-928">- Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-928">- Selection</span></span><br><span data-ttu-id="de7d0-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="de7d0-930">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="de7d0-930">Office 2013 on Windows</span></span><br><span data-ttu-id="de7d0-931">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="de7d0-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="de7d0-932">- Область задач</span><span class="sxs-lookup"><span data-stu-id="de7d0-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="de7d0-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="de7d0-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="de7d0-934">- Selection</span><span class="sxs-lookup"><span data-stu-id="de7d0-934">- Selection</span></span><br><span data-ttu-id="de7d0-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="de7d0-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="de7d0-936">См. также</span><span class="sxs-lookup"><span data-stu-id="de7d0-936">See also</span></span>

- [<span data-ttu-id="de7d0-937">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="de7d0-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="de7d0-938">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="de7d0-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="de7d0-939">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="de7d0-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="de7d0-940">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="de7d0-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="de7d0-941">Справочная документация по API</span><span class="sxs-lookup"><span data-stu-id="de7d0-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="de7d0-942">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="de7d0-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="de7d0-943">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="de7d0-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="de7d0-944">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="de7d0-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="de7d0-945">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="de7d0-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="de7d0-946">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="de7d0-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="de7d0-947">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="de7d0-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="de7d0-948">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="de7d0-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)