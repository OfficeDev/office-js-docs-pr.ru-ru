---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 06/13/2019
localization_priority: Priority
ms.openlocfilehash: 82c276c802cab66ae4f5443d0d556bc42ee57841
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128624"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="75d2e-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="75d2e-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="75d2e-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="75d2e-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="75d2e-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="75d2e-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="75d2e-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="75d2e-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="75d2e-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="75d2e-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="75d2e-108">Excel</span><span class="sxs-lookup"><span data-stu-id="75d2e-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="75d2e-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="75d2e-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="75d2e-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="75d2e-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="75d2e-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="75d2e-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="75d2e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="75d2e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="75d2e-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="75d2e-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-114">- TaskPane</span></span><br><span data-ttu-id="75d2e-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-115">
        - Content</span></span><br><span data-ttu-id="75d2e-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="75d2e-116">
        - Custom Functions</span></span><br><span data-ttu-id="75d2e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="75d2e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="75d2e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75d2e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75d2e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75d2e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75d2e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75d2e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75d2e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75d2e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75d2e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="75d2e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="75d2e-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-128">
        - BindingEvents</span></span><br><span data-ttu-id="75d2e-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-129">
        - CompressedFile</span></span><br><span data-ttu-id="75d2e-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-130">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-131">
        - File</span></span><br><span data-ttu-id="75d2e-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-132">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-134">
        - Selection</span></span><br><span data-ttu-id="75d2e-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-135">
        - Settings</span></span><br><span data-ttu-id="75d2e-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-136">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-137">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-138">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-140">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-140">Office on Windows</span></span><br><span data-ttu-id="75d2e-141">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-142">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-142">- TaskPane</span></span><br><span data-ttu-id="75d2e-143">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-143">
        - Content</span></span><br><span data-ttu-id="75d2e-144">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="75d2e-144">
        - Custom Functions</span></span><br><span data-ttu-id="75d2e-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="75d2e-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="75d2e-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75d2e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75d2e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75d2e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75d2e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75d2e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75d2e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75d2e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75d2e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="75d2e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="75d2e-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-156">
        - BindingEvents</span></span><br><span data-ttu-id="75d2e-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-157">
        - CompressedFile</span></span><br><span data-ttu-id="75d2e-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-158">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-159">
        - File</span></span><br><span data-ttu-id="75d2e-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-160">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-162">
        - Selection</span></span><br><span data-ttu-id="75d2e-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-163">
        - Settings</span></span><br><span data-ttu-id="75d2e-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-164">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-165">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-166">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-168">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-168">Office 2019 on Windows</span></span><br><span data-ttu-id="75d2e-169">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75d2e-170">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-170">- TaskPane</span></span><br><span data-ttu-id="75d2e-171">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-171">
        - Content</span></span><br><span data-ttu-id="75d2e-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="75d2e-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75d2e-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75d2e-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75d2e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75d2e-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75d2e-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75d2e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75d2e-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75d2e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="75d2e-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-182">- BindingEvents</span></span><br><span data-ttu-id="75d2e-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-183">
        - CompressedFile</span></span><br><span data-ttu-id="75d2e-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-184">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-185">
        - File</span></span><br><span data-ttu-id="75d2e-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-186">
        - ImageCoercion</span></span><br><span data-ttu-id="75d2e-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-187">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-189">
        - Selection</span></span><br><span data-ttu-id="75d2e-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-190">
        - Settings</span></span><br><span data-ttu-id="75d2e-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-191">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-192">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-193">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-195">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-195">Office 2016 on Windows</span></span><br><span data-ttu-id="75d2e-196">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75d2e-197">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-197">- TaskPane</span></span><br><span data-ttu-id="75d2e-198">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-198">
        - Content</span></span></td>
    <td><span data-ttu-id="75d2e-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75d2e-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="75d2e-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="75d2e-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-201">- BindingEvents</span></span><br><span data-ttu-id="75d2e-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-202">
        - CompressedFile</span></span><br><span data-ttu-id="75d2e-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-203">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-204">
        - File</span></span><br><span data-ttu-id="75d2e-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-205">
        - ImageCoercion</span></span><br><span data-ttu-id="75d2e-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-206">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-208">
        - Selection</span></span><br><span data-ttu-id="75d2e-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-209">
        - Settings</span></span><br><span data-ttu-id="75d2e-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-210">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-211">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-212">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-214">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-214">Office 2013 on Windows</span></span><br><span data-ttu-id="75d2e-215">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75d2e-216">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-216">
        - TaskPane</span></span><br><span data-ttu-id="75d2e-217">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="75d2e-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75d2e-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="75d2e-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-219">
        - BindingEvents</span></span><br><span data-ttu-id="75d2e-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-220">
        - CompressedFile</span></span><br><span data-ttu-id="75d2e-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-221">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-222">
        - File</span></span><br><span data-ttu-id="75d2e-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-223">
        - ImageCoercion</span></span><br><span data-ttu-id="75d2e-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-224">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-226">
        - Selection</span></span><br><span data-ttu-id="75d2e-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-227">
        - Settings</span></span><br><span data-ttu-id="75d2e-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-228">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-229">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-230">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-232">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="75d2e-232">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="75d2e-233">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-233">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="75d2e-234">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-234">- TaskPane</span></span><br><span data-ttu-id="75d2e-235">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-235">
        - Content</span></span><br><span data-ttu-id="75d2e-236">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="75d2e-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="75d2e-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75d2e-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75d2e-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75d2e-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75d2e-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75d2e-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75d2e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75d2e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75d2e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="75d2e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="75d2e-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-247">- BindingEvents</span></span><br><span data-ttu-id="75d2e-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-248">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-249">
        - File</span></span><br><span data-ttu-id="75d2e-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-250">
        - ImageCoercion</span></span><br><span data-ttu-id="75d2e-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-251">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-253">
        - Selection</span></span><br><span data-ttu-id="75d2e-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-254">
        - Settings</span></span><br><span data-ttu-id="75d2e-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-255">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-256">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-257">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-259">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-259">Office apps on Mac</span></span><br><span data-ttu-id="75d2e-260">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-260">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="75d2e-261">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-261">- TaskPane</span></span><br><span data-ttu-id="75d2e-262">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-262">
        - Content</span></span><br><span data-ttu-id="75d2e-263">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="75d2e-263">
        - Custom Functions</span></span><br><span data-ttu-id="75d2e-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="75d2e-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75d2e-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75d2e-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75d2e-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75d2e-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75d2e-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75d2e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75d2e-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75d2e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="75d2e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="75d2e-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-275">- BindingEvents</span></span><br><span data-ttu-id="75d2e-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-276">
        - CompressedFile</span></span><br><span data-ttu-id="75d2e-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-277">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-278">
        - File</span></span><br><span data-ttu-id="75d2e-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-279">
        - ImageCoercion</span></span><br><span data-ttu-id="75d2e-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-280">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-282">
        - PdfFile</span></span><br><span data-ttu-id="75d2e-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-283">
        - Selection</span></span><br><span data-ttu-id="75d2e-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-284">
        - Settings</span></span><br><span data-ttu-id="75d2e-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-285">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-286">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-287">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-289">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-289">Office 2019 for Mac</span></span><br><span data-ttu-id="75d2e-290">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75d2e-291">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-291">- TaskPane</span></span><br><span data-ttu-id="75d2e-292">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-292">
        - Content</span></span><br><span data-ttu-id="75d2e-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="75d2e-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75d2e-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75d2e-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75d2e-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75d2e-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75d2e-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75d2e-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75d2e-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75d2e-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="75d2e-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-303">- BindingEvents</span></span><br><span data-ttu-id="75d2e-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-304">
        - CompressedFile</span></span><br><span data-ttu-id="75d2e-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-305">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-306">
        - File</span></span><br><span data-ttu-id="75d2e-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-307">
        - ImageCoercion</span></span><br><span data-ttu-id="75d2e-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-308">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-310">
        - PdfFile</span></span><br><span data-ttu-id="75d2e-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-311">
        - Selection</span></span><br><span data-ttu-id="75d2e-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-312">
        - Settings</span></span><br><span data-ttu-id="75d2e-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-313">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-314">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-315">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-317">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-317">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="75d2e-318">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75d2e-319">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-319">- TaskPane</span></span><br><span data-ttu-id="75d2e-320">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-320">
        - Content</span></span></td>
    <td><span data-ttu-id="75d2e-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75d2e-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="75d2e-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="75d2e-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-323">- BindingEvents</span></span><br><span data-ttu-id="75d2e-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-324">
        - CompressedFile</span></span><br><span data-ttu-id="75d2e-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-325">
        - DocumentEvents</span></span><br><span data-ttu-id="75d2e-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-326">
        - File</span></span><br><span data-ttu-id="75d2e-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-327">
        - ImageCoercion</span></span><br><span data-ttu-id="75d2e-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-328">
        - MatrixBindings</span></span><br><span data-ttu-id="75d2e-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-330">
        - PdfFile</span></span><br><span data-ttu-id="75d2e-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-331">
        - Selection</span></span><br><span data-ttu-id="75d2e-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-332">
        - Settings</span></span><br><span data-ttu-id="75d2e-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-333">
        - TableBindings</span></span><br><span data-ttu-id="75d2e-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-334">
        - TableCoercion</span></span><br><span data-ttu-id="75d2e-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-335">
        - TextBindings</span></span><br><span data-ttu-id="75d2e-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="75d2e-337">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="75d2e-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="75d2e-338">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="75d2e-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="75d2e-339">Платформа</span><span class="sxs-lookup"><span data-stu-id="75d2e-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="75d2e-340">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="75d2e-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="75d2e-341">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="75d2e-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="75d2e-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="75d2e-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-343">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="75d2e-343">Office on the web</span></span></td>
    <td><span data-ttu-id="75d2e-344">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="75d2e-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="75d2e-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-346">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-346">Office on Windows</span></span><br><span data-ttu-id="75d2e-347">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-347">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="75d2e-348">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="75d2e-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="75d2e-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-350">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-350">Office for Mac</span></span><br><span data-ttu-id="75d2e-351">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="75d2e-352">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="75d2e-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="75d2e-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="75d2e-354">Outlook</span><span class="sxs-lookup"><span data-stu-id="75d2e-354">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75d2e-355">Платформа</span><span class="sxs-lookup"><span data-stu-id="75d2e-355">Platform</span></span></th>
    <th><span data-ttu-id="75d2e-356">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="75d2e-356">Extension points</span></span></th>
    <th><span data-ttu-id="75d2e-357">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="75d2e-357">API requirement sets</span></span></th>
    <th><span data-ttu-id="75d2e-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="75d2e-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-359">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="75d2e-359">Office on the web</span></span><br><span data-ttu-id="75d2e-360">(новый)</span><span class="sxs-lookup"><span data-stu-id="75d2e-360">New</span></span></td>
    <td> <span data-ttu-id="75d2e-361">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-361">- Mail Read</span></span><br><span data-ttu-id="75d2e-362">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-362">
      - Mail Compose</span></span><br><span data-ttu-id="75d2e-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75d2e-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="75d2e-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="75d2e-371">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-372">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="75d2e-372">Office on the web</span></span><br><span data-ttu-id="75d2e-373">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="75d2e-373">(classic)</span></span></td>
    <td> <span data-ttu-id="75d2e-374">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-374">- Mail Read</span></span><br><span data-ttu-id="75d2e-375">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-375">
      - Mail Compose</span></span><br><span data-ttu-id="75d2e-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75d2e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="75d2e-383">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-384">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-384">Office on Windows</span></span><br><span data-ttu-id="75d2e-385">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-385">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-386">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-386">- Mail Read</span></span><br><span data-ttu-id="75d2e-387">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-387">
      - Mail Compose</span></span><br><span data-ttu-id="75d2e-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="75d2e-389">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="75d2e-389">
      - Modules</span></span></td>
    <td> <span data-ttu-id="75d2e-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75d2e-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="75d2e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="75d2e-397">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-397">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-398">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-398">Office 2019 on Windows</span></span><br><span data-ttu-id="75d2e-399">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-399">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-400">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-400">- Mail Read</span></span><br><span data-ttu-id="75d2e-401">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-401">
      - Mail Compose</span></span><br><span data-ttu-id="75d2e-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="75d2e-403">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="75d2e-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="75d2e-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75d2e-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="75d2e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="75d2e-411">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-411">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-412">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-412">Office 2016 on Windows</span></span><br><span data-ttu-id="75d2e-413">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-413">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-414">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-414">- Mail Read</span></span><br><span data-ttu-id="75d2e-415">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-415">
      - Mail Compose</span></span><br><span data-ttu-id="75d2e-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="75d2e-417">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="75d2e-417">
      - Modules</span></span></td>
    <td> <span data-ttu-id="75d2e-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="75d2e-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="75d2e-422">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-423">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-423">Office 2013 on Windows</span></span><br><span data-ttu-id="75d2e-424">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-424">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-425">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-425">- Mail Read</span></span><br><span data-ttu-id="75d2e-426">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-426">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="75d2e-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="75d2e-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="75d2e-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="75d2e-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="75d2e-431">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-432">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="75d2e-432">Office apps on iOS</span></span><br><span data-ttu-id="75d2e-433">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-433">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-434">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-434">- Mail Read</span></span><br><span data-ttu-id="75d2e-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="75d2e-441">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-442">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-442">Office apps on Mac</span></span><br><span data-ttu-id="75d2e-443">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-443">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-444">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-444">- Mail Read</span></span><br><span data-ttu-id="75d2e-445">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-445">
      - Mail Compose</span></span><br><span data-ttu-id="75d2e-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75d2e-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="75d2e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="75d2e-454">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-454">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-455">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-455">Office 2019 for Mac</span></span><br><span data-ttu-id="75d2e-456">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-456">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-457">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-457">- Mail Read</span></span><br><span data-ttu-id="75d2e-458">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-458">
      - Mail Compose</span></span><br><span data-ttu-id="75d2e-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75d2e-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="75d2e-466">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-467">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-467">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="75d2e-468">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-468">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-469">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-469">- Mail Read</span></span><br><span data-ttu-id="75d2e-470">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-470">
      - Mail Compose</span></span><br><span data-ttu-id="75d2e-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75d2e-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="75d2e-478">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-479">Office для Android</span><span class="sxs-lookup"><span data-stu-id="75d2e-479">Office apps on Android</span></span><br><span data-ttu-id="75d2e-480">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-480">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-481">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="75d2e-481">- Mail Read</span></span><br><span data-ttu-id="75d2e-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75d2e-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75d2e-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75d2e-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75d2e-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="75d2e-488">Недоступно</span><span class="sxs-lookup"><span data-stu-id="75d2e-488">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="75d2e-489">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="75d2e-489">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="75d2e-490">Word</span><span class="sxs-lookup"><span data-stu-id="75d2e-490">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75d2e-491">Платформа</span><span class="sxs-lookup"><span data-stu-id="75d2e-491">Platform</span></span></th>
    <th><span data-ttu-id="75d2e-492">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="75d2e-492">Extension points</span></span></th>
    <th><span data-ttu-id="75d2e-493">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="75d2e-493">API requirement sets</span></span></th>
    <th><span data-ttu-id="75d2e-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="75d2e-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-495">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="75d2e-495">Office on the web</span></span></td>
    <td> <span data-ttu-id="75d2e-496">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-496">- TaskPane</span></span><br><span data-ttu-id="75d2e-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="75d2e-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="75d2e-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="75d2e-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-502">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-502">- BindingEvents</span></span><br><span data-ttu-id="75d2e-503">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-503">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-504">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-504">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-505">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-505">
         - File</span></span><br><span data-ttu-id="75d2e-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-506">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-507">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-508">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-508">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-509">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-509">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-510">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-510">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-511">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-511">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-512">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-512">
         - Selection</span></span><br><span data-ttu-id="75d2e-513">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-513">
         - Settings</span></span><br><span data-ttu-id="75d2e-514">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-514">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-515">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-515">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-516">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-516">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-517">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-517">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-518">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-518">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-519">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-519">Office on Windows</span></span><br><span data-ttu-id="75d2e-520">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-520">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-521">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-521">- TaskPane</span></span><br><span data-ttu-id="75d2e-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="75d2e-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="75d2e-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="75d2e-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-527">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-527">- BindingEvents</span></span><br><span data-ttu-id="75d2e-528">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-528">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-529">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-529">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-530">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-530">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-531">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-531">
         - File</span></span><br><span data-ttu-id="75d2e-532">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-532">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-533">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-533">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-534">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-534">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-535">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-535">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-536">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-536">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-537">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-537">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-538">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-538">
         - Selection</span></span><br><span data-ttu-id="75d2e-539">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-539">
         - Settings</span></span><br><span data-ttu-id="75d2e-540">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-540">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-541">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-541">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-542">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-542">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-543">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-543">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-544">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-544">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-545">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-545">Office 2019 on Windows</span></span><br><span data-ttu-id="75d2e-546">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-546">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-547">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-547">- TaskPane</span></span><br><span data-ttu-id="75d2e-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="75d2e-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="75d2e-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="75d2e-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-553">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-553">- BindingEvents</span></span><br><span data-ttu-id="75d2e-554">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-554">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-555">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-555">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-556">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-557">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-557">
         - File</span></span><br><span data-ttu-id="75d2e-558">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-558">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-559">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-559">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-560">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-560">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-561">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-561">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-562">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-562">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-563">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-563">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-564">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-564">
         - Selection</span></span><br><span data-ttu-id="75d2e-565">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-565">
         - Settings</span></span><br><span data-ttu-id="75d2e-566">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-566">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-567">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-567">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-568">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-568">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-569">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-569">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-570">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-570">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-571">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-571">Office 2016 on Windows</span></span><br><span data-ttu-id="75d2e-572">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-572">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-573">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-573">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="75d2e-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="75d2e-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="75d2e-576">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-576">- BindingEvents</span></span><br><span data-ttu-id="75d2e-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-577">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-578">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-578">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-579">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-580">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-580">
         - File</span></span><br><span data-ttu-id="75d2e-581">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-581">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-582">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-583">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-583">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-584">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-584">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-585">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-585">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-586">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-587">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-587">
         - Selection</span></span><br><span data-ttu-id="75d2e-588">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-588">
         - Settings</span></span><br><span data-ttu-id="75d2e-589">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-589">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-590">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-590">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-591">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-591">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-592">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-592">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-593">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-593">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-594">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-594">Office 2013 on Windows</span></span><br><span data-ttu-id="75d2e-595">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-595">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-596">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-596">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75d2e-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="75d2e-598">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-598">- BindingEvents</span></span><br><span data-ttu-id="75d2e-599">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-599">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-600">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-600">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-601">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-601">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-602">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-602">
         - File</span></span><br><span data-ttu-id="75d2e-603">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-603">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-604">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-604">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-605">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-605">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-606">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-606">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-607">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-607">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-608">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-609">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-609">
         - Selection</span></span><br><span data-ttu-id="75d2e-610">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-610">
         - Settings</span></span><br><span data-ttu-id="75d2e-611">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-611">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-612">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-612">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-613">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-613">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-614">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-615">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-615">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-616">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="75d2e-616">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="75d2e-617">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-617">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-618">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-618">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="75d2e-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="75d2e-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="75d2e-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="75d2e-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="75d2e-623">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-623">- BindingEvents</span></span><br><span data-ttu-id="75d2e-624">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-624">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-625">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-625">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-626">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-627">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-627">
         - File</span></span><br><span data-ttu-id="75d2e-628">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-628">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-629">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-629">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-630">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-630">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-631">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-631">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-632">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-632">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-633">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-633">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-634">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-634">
         - Selection</span></span><br><span data-ttu-id="75d2e-635">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-635">
         - Settings</span></span><br><span data-ttu-id="75d2e-636">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-636">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-637">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-637">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-638">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-638">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-639">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-640">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-640">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-641">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-641">Office apps on Mac</span></span><br><span data-ttu-id="75d2e-642">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-642">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-643">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-643">- TaskPane</span></span><br><span data-ttu-id="75d2e-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="75d2e-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="75d2e-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="75d2e-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="75d2e-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="75d2e-649">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-649">- BindingEvents</span></span><br><span data-ttu-id="75d2e-650">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-650">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-651">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-651">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-652">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-652">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-653">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-653">
         - File</span></span><br><span data-ttu-id="75d2e-654">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-654">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-655">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-655">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-656">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-656">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-657">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-657">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-658">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-658">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-659">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-659">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-660">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-660">
         - Selection</span></span><br><span data-ttu-id="75d2e-661">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-661">
         - Settings</span></span><br><span data-ttu-id="75d2e-662">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-662">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-663">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-663">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-664">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-664">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-665">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-665">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-666">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-666">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-667">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-667">Office 2019 for Mac</span></span><br><span data-ttu-id="75d2e-668">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-668">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-669">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-669">- TaskPane</span></span><br><span data-ttu-id="75d2e-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="75d2e-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="75d2e-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="75d2e-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="75d2e-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="75d2e-675">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-675">- BindingEvents</span></span><br><span data-ttu-id="75d2e-676">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-676">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-677">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-677">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-678">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-678">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-679">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-679">
         - File</span></span><br><span data-ttu-id="75d2e-680">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-680">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-681">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-681">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-682">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-682">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-683">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-683">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-684">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-684">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-685">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-686">
         - Selection</span></span><br><span data-ttu-id="75d2e-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-687">
         - Settings</span></span><br><span data-ttu-id="75d2e-688">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-688">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-689">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-689">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-690">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-690">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-691">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-691">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-692">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-692">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-693">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-693">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="75d2e-694">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-694">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-695">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-695">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="75d2e-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="75d2e-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="75d2e-698">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-698">- BindingEvents</span></span><br><span data-ttu-id="75d2e-699">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-699">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-700">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75d2e-700">
         - CustomXmlParts</span></span><br><span data-ttu-id="75d2e-701">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-701">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-702">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="75d2e-702">
         - File</span></span><br><span data-ttu-id="75d2e-703">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-703">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-704">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-704">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-705">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-705">
         - MatrixBindings</span></span><br><span data-ttu-id="75d2e-706">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-706">
         - MatrixCoercion</span></span><br><span data-ttu-id="75d2e-707">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-707">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75d2e-708">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-708">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-709">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-709">
         - Selection</span></span><br><span data-ttu-id="75d2e-710">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="75d2e-710">
         - Settings</span></span><br><span data-ttu-id="75d2e-711">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-711">
         - TableBindings</span></span><br><span data-ttu-id="75d2e-712">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-712">
         - TableCoercion</span></span><br><span data-ttu-id="75d2e-713">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75d2e-713">
         - TextBindings</span></span><br><span data-ttu-id="75d2e-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-714">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-715">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-715">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="75d2e-716">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="75d2e-716">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="75d2e-717">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="75d2e-717">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75d2e-718">Платформа</span><span class="sxs-lookup"><span data-stu-id="75d2e-718">Platform</span></span></th>
    <th><span data-ttu-id="75d2e-719">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="75d2e-719">Extension points</span></span></th>
    <th><span data-ttu-id="75d2e-720">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="75d2e-720">API requirement sets</span></span></th>
    <th><span data-ttu-id="75d2e-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="75d2e-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-722">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="75d2e-722">Office on the web</span></span></td>
    <td> <span data-ttu-id="75d2e-723">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-723">- Content</span></span><br><span data-ttu-id="75d2e-724">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-724">
         - TaskPane</span></span><br><span data-ttu-id="75d2e-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-727">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-727">- ActiveView</span></span><br><span data-ttu-id="75d2e-728">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-728">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-729">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-729">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-730">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-730">
         - File</span></span><br><span data-ttu-id="75d2e-731">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-731">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-732">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-732">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-733">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-733">
         - Selection</span></span><br><span data-ttu-id="75d2e-734">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-734">
         - Settings</span></span><br><span data-ttu-id="75d2e-735">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-735">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-736">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-736">Office on Windows</span></span><br><span data-ttu-id="75d2e-737">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-737">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-738">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-738">- Content</span></span><br><span data-ttu-id="75d2e-739">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-739">
         - TaskPane</span></span><br><span data-ttu-id="75d2e-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-742">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-742">- ActiveView</span></span><br><span data-ttu-id="75d2e-743">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-743">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-744">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-744">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-745">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-745">
         - File</span></span><br><span data-ttu-id="75d2e-746">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-746">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-747">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-747">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-748">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-748">
         - Selection</span></span><br><span data-ttu-id="75d2e-749">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-749">
         - Settings</span></span><br><span data-ttu-id="75d2e-750">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-750">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-751">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-751">Office 2019 on Windows</span></span><br><span data-ttu-id="75d2e-752">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-752">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-753">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-753">- Content</span></span><br><span data-ttu-id="75d2e-754">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-754">
         - TaskPane</span></span><br><span data-ttu-id="75d2e-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-757">- ActiveView</span></span><br><span data-ttu-id="75d2e-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-758">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-759">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-760">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-760">
         - File</span></span><br><span data-ttu-id="75d2e-761">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-761">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-762">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-762">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-763">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-763">
         - Selection</span></span><br><span data-ttu-id="75d2e-764">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-764">
         - Settings</span></span><br><span data-ttu-id="75d2e-765">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-765">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-766">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-766">Office 2016 on Windows</span></span><br><span data-ttu-id="75d2e-767">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-767">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-768">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-768">- Content</span></span><br><span data-ttu-id="75d2e-769">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-769">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75d2e-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="75d2e-771">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-771">- ActiveView</span></span><br><span data-ttu-id="75d2e-772">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-772">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-773">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-773">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-774">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-774">
         - File</span></span><br><span data-ttu-id="75d2e-775">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-775">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-776">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-777">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-777">
         - Selection</span></span><br><span data-ttu-id="75d2e-778">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-778">
         - Settings</span></span><br><span data-ttu-id="75d2e-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-780">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-780">Office 2013 on Windows</span></span><br><span data-ttu-id="75d2e-781">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-782">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-782">- Content</span></span><br><span data-ttu-id="75d2e-783">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-783">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="75d2e-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75d2e-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="75d2e-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-785">- ActiveView</span></span><br><span data-ttu-id="75d2e-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-786">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-787">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-788">
         - File</span></span><br><span data-ttu-id="75d2e-789">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-789">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-790">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-791">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-791">
         - Selection</span></span><br><span data-ttu-id="75d2e-792">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-792">
         - Settings</span></span><br><span data-ttu-id="75d2e-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-794">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="75d2e-794">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="75d2e-795">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-795">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-796">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-796">- Content</span></span><br><span data-ttu-id="75d2e-797">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-797">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-799">- ActiveView</span></span><br><span data-ttu-id="75d2e-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-800">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-801">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-802">
         - File</span></span><br><span data-ttu-id="75d2e-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-803">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-804">
         - Selection</span></span><br><span data-ttu-id="75d2e-805">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-805">
         - Settings</span></span><br><span data-ttu-id="75d2e-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-806">
         - TextCoercion</span></span><br><span data-ttu-id="75d2e-807">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-807">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-808">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-808">Office apps on Mac</span></span><br><span data-ttu-id="75d2e-809">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="75d2e-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75d2e-810">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-810">- Content</span></span><br><span data-ttu-id="75d2e-811">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-811">
         - TaskPane</span></span><br><span data-ttu-id="75d2e-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-814">- ActiveView</span></span><br><span data-ttu-id="75d2e-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-815">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-816">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-817">
         - File</span></span><br><span data-ttu-id="75d2e-818">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-818">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-819">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-820">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-820">
         - Selection</span></span><br><span data-ttu-id="75d2e-821">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-821">
         - Settings</span></span><br><span data-ttu-id="75d2e-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-823">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-823">Office 2019 for Mac</span></span><br><span data-ttu-id="75d2e-824">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-824">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-825">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-825">- Content</span></span><br><span data-ttu-id="75d2e-826">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-826">
         - TaskPane</span></span><br><span data-ttu-id="75d2e-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-829">- ActiveView</span></span><br><span data-ttu-id="75d2e-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-830">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-831">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-832">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-832">
         - File</span></span><br><span data-ttu-id="75d2e-833">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-833">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-834">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-835">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-835">
         - Selection</span></span><br><span data-ttu-id="75d2e-836">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-836">
         - Settings</span></span><br><span data-ttu-id="75d2e-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-838">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-838">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="75d2e-839">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-840">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-840">- Content</span></span><br><span data-ttu-id="75d2e-841">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-841">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75d2e-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="75d2e-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75d2e-843">- ActiveView</span></span><br><span data-ttu-id="75d2e-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-844">
         - CompressedFile</span></span><br><span data-ttu-id="75d2e-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-845">
         - DocumentEvents</span></span><br><span data-ttu-id="75d2e-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="75d2e-846">
         - File</span></span><br><span data-ttu-id="75d2e-847">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-847">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75d2e-848">
         - PdfFile</span></span><br><span data-ttu-id="75d2e-849">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-849">
         - Selection</span></span><br><span data-ttu-id="75d2e-850">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-850">
         - Settings</span></span><br><span data-ttu-id="75d2e-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-851">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="75d2e-852">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="75d2e-852">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="75d2e-853">OneNote</span><span class="sxs-lookup"><span data-stu-id="75d2e-853">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75d2e-854">Платформа</span><span class="sxs-lookup"><span data-stu-id="75d2e-854">Platform</span></span></th>
    <th><span data-ttu-id="75d2e-855">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="75d2e-855">Extension points</span></span></th>
    <th><span data-ttu-id="75d2e-856">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="75d2e-856">API requirement sets</span></span></th>
    <th><span data-ttu-id="75d2e-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="75d2e-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-858">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="75d2e-858">Office on the web</span></span></td>
    <td> <span data-ttu-id="75d2e-859">- Контент</span><span class="sxs-lookup"><span data-stu-id="75d2e-859">- Content</span></span><br><span data-ttu-id="75d2e-860">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-860">
         - TaskPane</span></span><br><span data-ttu-id="75d2e-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75d2e-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="75d2e-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-864">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75d2e-864">- DocumentEvents</span></span><br><span data-ttu-id="75d2e-865">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-865">
         - HtmlCoercion</span></span><br><span data-ttu-id="75d2e-866">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-866">
         - ImageCoercion</span></span><br><span data-ttu-id="75d2e-867">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="75d2e-867">
         - Settings</span></span><br><span data-ttu-id="75d2e-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="75d2e-869">Project</span><span class="sxs-lookup"><span data-stu-id="75d2e-869">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75d2e-870">Платформа</span><span class="sxs-lookup"><span data-stu-id="75d2e-870">Platform</span></span></th>
    <th><span data-ttu-id="75d2e-871">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="75d2e-871">Extension points</span></span></th>
    <th><span data-ttu-id="75d2e-872">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="75d2e-872">API requirement sets</span></span></th>
    <th><span data-ttu-id="75d2e-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="75d2e-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-874">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-874">Office 2019 on Windows</span></span><br><span data-ttu-id="75d2e-875">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-875">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-876">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-876">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-878">- Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-878">- Selection</span></span><br><span data-ttu-id="75d2e-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-879">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-880">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-880">Office 2016 on Windows</span></span><br><span data-ttu-id="75d2e-881">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-881">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-882">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-882">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-884">- Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-884">- Selection</span></span><br><span data-ttu-id="75d2e-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-885">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75d2e-886">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="75d2e-886">Office 2013 on Windows</span></span><br><span data-ttu-id="75d2e-887">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="75d2e-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75d2e-888">- Область задач</span><span class="sxs-lookup"><span data-stu-id="75d2e-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75d2e-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75d2e-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75d2e-890">- Selection</span><span class="sxs-lookup"><span data-stu-id="75d2e-890">- Selection</span></span><br><span data-ttu-id="75d2e-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75d2e-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="75d2e-892">См. также</span><span class="sxs-lookup"><span data-stu-id="75d2e-892">See also</span></span>

- [<span data-ttu-id="75d2e-893">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="75d2e-893">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="75d2e-894">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="75d2e-894">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="75d2e-895">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="75d2e-895">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="75d2e-896">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="75d2e-896">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="75d2e-897">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="75d2e-897">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="75d2e-898">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="75d2e-898">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="75d2e-899">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="75d2e-899">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="75d2e-900">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="75d2e-900">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="75d2e-901">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="75d2e-901">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="75d2e-902">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="75d2e-902">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="75d2e-903">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="75d2e-903">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
