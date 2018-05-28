---
title: ??????????? ??????? ?????????? ? ???????? ??? ????????? Office
description: ?????????????? ?????? ?????????? ??? Excel, Word, Outlook, PowerPoint ? OneNote.
ms.date: 03/23/2018
ms.openlocfilehash: f50ab7e5312702eb25fbb2c8a25291c5ff5027a7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="cd205-103">??????????? ??????? ?????????? ? ???????? ??? ????????? Office</span><span class="sxs-lookup"><span data-stu-id="cd205-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="cd205-104">?????? ?????????? Office ????? ???????? ?? ???????? ?????????? Office, ?????? ??????????, ???????? ??? ?????? API.</span><span class="sxs-lookup"><span data-stu-id="cd205-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="cd205-105">? ???????? ???? ???????????? ???????? ? ????????? ?????????, ?????? ??????????, ??????? ???????????? ????????? API ? ??????????? ??????? ???????????? ????????? API, ??????? ? ????????? ????? ?????????????? ??? ???? ?????????? Office.</span><span class="sxs-lookup"><span data-stu-id="cd205-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="cd205-106">?????? \* (?????????) ? ?????? ??????? ?????????, ??? ????????? ????? ????????.</span><span class="sxs-lookup"><span data-stu-id="cd205-106">If a table cell contains an asterisk ( \* ), that means we?re working on it.</span></span> <span data-ttu-id="cd205-107">? ???????? ?????????? ??? Project ? Access ????? ???????????? ? ?????? [??????????? ?????? ???????????? ????????? ??? Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="cd205-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="cd205-p103">????? ?????? ??? ?????? Office 2016, ?????????????? ? ??????? MSI, ? 16.0.4266.1001. ??? ?????? ???????? ?????? ????? ???????????? ????????? ExcelApi 1.1, WordApi 1.1 ? ??????????? ?????? ???????????? ????????? API.</span><span class="sxs-lookup"><span data-stu-id="cd205-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="cd205-110">Excel</span><span class="sxs-lookup"><span data-stu-id="cd205-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="cd205-111">?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="cd205-112">????? ??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="cd205-113">?????? ???????????? ????????? API</span><span class="sxs-lookup"><span data-stu-id="cd205-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="cd205-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></span><span class="sxs-lookup"><span data-stu-id="cd205-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="cd205-115">Office Online</span></span></td>
    <td> <span data-ttu-id="cd205-116">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-116">- Taskpane</span></span><br><span data-ttu-id="cd205-117">
        - ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-117">
        - Content</span></span><br><span data-ttu-id="cd205-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ??????????</a>
    </span><span class="sxs-lookup"><span data-stu-id="cd205-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="cd205-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cd205-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cd205-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cd205-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cd205-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="cd205-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cd205-124">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-124">
        -BindingEvents</span></span><br><span data-ttu-id="cd205-125">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-125">
        -DocumentEvents</span></span><br><span data-ttu-id="cd205-126">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-126">
        -MatrixBindings</span></span><br><span data-ttu-id="cd205-127">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-127">
        -MatrixCoercion</span></span><br><span data-ttu-id="cd205-128">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-128">
        -TableBindings</span></span><br><span data-ttu-id="cd205-129">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-129">
        -TableCoercion</span></span><br><span data-ttu-id="cd205-130">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-130">
        -TextBindings</span></span><br><span data-ttu-id="cd205-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-131">
        -CompressedFile</span></span><br><span data-ttu-id="cd205-132">
        - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-132">
        - Settings</span></span><br><span data-ttu-id="cd205-133">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-133">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-134">Office 2013 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-134">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="cd205-135">
        - ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-135">
        - Taskpane</span></span><br><span data-ttu-id="cd205-136">
        - ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-136">
        - Content</span></span></td>
    <td>  <span data-ttu-id="cd205-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cd205-138">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-138">
        -BindingEvents</span></span><br><span data-ttu-id="cd205-139">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-139">
        -DocumentEvents</span></span><br><span data-ttu-id="cd205-140">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-140">
        -MatrixBindings</span></span><br><span data-ttu-id="cd205-141">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-141">
        -MatrixCoercion</span></span><br><span data-ttu-id="cd205-142">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-142">
        -TableBindings</span></span><br><span data-ttu-id="cd205-143">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-143">
        -TableCoercion</span></span><br><span data-ttu-id="cd205-144">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-144">
        -TextBindings</span></span><br><span data-ttu-id="cd205-145">
        - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-145">
        - Settings</span></span><br><span data-ttu-id="cd205-146">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-146">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-147">Office 2016 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-147">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="cd205-148">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-148">- Taskpane</span></span><br><span data-ttu-id="cd205-149">
        - ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-149">
        - Content</span></span><br><span data-ttu-id="cd205-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="cd205-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cd205-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cd205-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cd205-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cd205-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="cd205-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cd205-156">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-156">-BindingEvents</span></span><br><span data-ttu-id="cd205-157">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-157">
        -DocumentEvents</span></span><br><span data-ttu-id="cd205-158">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-158">
        -MatrixBindings</span></span><br><span data-ttu-id="cd205-159">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-159">
        -MatrixCoercion</span></span><br><span data-ttu-id="cd205-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-160">
        -TableBindings</span></span><br><span data-ttu-id="cd205-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-161">
        -TableCoercion</span></span><br><span data-ttu-id="cd205-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-162">
        -TextBindings</span></span><br><span data-ttu-id="cd205-163">
        - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-163">
        - Settings</span></span><br><span data-ttu-id="cd205-164">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-164">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-165">Office ??? iOS</span><span class="sxs-lookup"><span data-stu-id="cd205-165">Office for iOS</span></span></td>
    <td><span data-ttu-id="cd205-166">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-166">- Taskpane</span></span><br><span data-ttu-id="cd205-167">
        - ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-167">
        - Content</span></span></td>
    <td><span data-ttu-id="cd205-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cd205-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cd205-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cd205-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cd205-172">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-172">-BindingEvents</span></span><br><span data-ttu-id="cd205-173">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-173">
        -DocumentEvents</span></span><br><span data-ttu-id="cd205-174">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-174">
        -MatrixBindings</span></span><br><span data-ttu-id="cd205-175">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-175">
        -MatrixCoercion</span></span><br><span data-ttu-id="cd205-176">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-176">
        -TableBindings</span></span><br><span data-ttu-id="cd205-177">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-177">
        -TableCoercion</span></span><br><span data-ttu-id="cd205-178">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-178">
        -TextBindings</span></span><br><span data-ttu-id="cd205-179">
        - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-179">
        - Settings</span></span><br><span data-ttu-id="cd205-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-181">Office 2016 ??? Mac</span><span class="sxs-lookup"><span data-stu-id="cd205-181">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="cd205-182">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-182">- Taskpane</span></span><br><span data-ttu-id="cd205-183">
        - ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-183">
        - Content</span></span><br><span data-ttu-id="cd205-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="cd205-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cd205-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cd205-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cd205-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cd205-189">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-189">-BindingEvents</span></span><br><span data-ttu-id="cd205-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-190">
        -DocumentEvents</span></span><br><span data-ttu-id="cd205-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-191">
        -MatrixBindings</span></span><br><span data-ttu-id="cd205-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-192">
        -MatrixCoercion</span></span><br><span data-ttu-id="cd205-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-193">
        -TableBindings</span></span><br><span data-ttu-id="cd205-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-194">
        -TableCoercion</span></span><br><span data-ttu-id="cd205-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-195">
        -TextBindings</span></span><br><span data-ttu-id="cd205-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-196">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="cd205-197">Outlook</span><span class="sxs-lookup"><span data-stu-id="cd205-197">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cd205-198">?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-198">Platform</span></span></th>
    <th><span data-ttu-id="cd205-199">????? ??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-199">Extension points</span></span></th> 
    <th><span data-ttu-id="cd205-200">?????? ???????????? ????????? API</span><span class="sxs-lookup"><span data-stu-id="cd205-200">API requirement sets</span></span></th> 
    <th><span data-ttu-id="cd205-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></span><span class="sxs-lookup"><span data-stu-id="cd205-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-202">Office Online</span><span class="sxs-lookup"><span data-stu-id="cd205-202">Office Online</span></span></td>
    <td> <span data-ttu-id="cd205-203">- ?????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-203">- Mail Read</span></span><br><span data-ttu-id="cd205-204">
      - ???????? ????????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-204">
      - Mail Compose</span></span><br><span data-ttu-id="cd205-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cd205-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cd205-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cd205-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cd205-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cd205-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cd205-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cd205-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cd205-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="cd205-212">??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-212">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-213">Office 2013 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-213">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="cd205-214">- ?????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-214">- Mail Read</span></span><br><span data-ttu-id="cd205-215">
      - ???????? ????????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-215">
      - Mail Compose</span></span><br><span data-ttu-id="cd205-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cd205-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cd205-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cd205-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cd205-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="cd205-221">??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-221">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-222">Office 2016 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-222">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="cd205-223">- ?????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-223">- Mail Read</span></span><br><span data-ttu-id="cd205-224">
      - ???????? ????????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-224">
      - Mail Compose</span></span><br><span data-ttu-id="cd205-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="cd205-226">
      - ??????</span><span class="sxs-lookup"><span data-stu-id="cd205-226">
      - Modules</span></span></td>
    <td> <span data-ttu-id="cd205-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cd205-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cd205-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cd205-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cd205-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cd205-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cd205-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cd205-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cd205-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="cd205-233">??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-233">not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-234">Office ??? iOS</span><span class="sxs-lookup"><span data-stu-id="cd205-234">Office for iOS</span></span></td>
    <td> <span data-ttu-id="cd205-235">- ?????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-235">- Mail Read</span></span><br><span data-ttu-id="cd205-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cd205-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cd205-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cd205-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cd205-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cd205-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cd205-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="cd205-242">??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-242">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-243">Office 2016 ??? Mac</span><span class="sxs-lookup"><span data-stu-id="cd205-243">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="cd205-244">- ?????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-244">- Mail Read</span></span><br><span data-ttu-id="cd205-245">
      - ???????? ????????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-245">
      - Mail Compose</span></span><br><span data-ttu-id="cd205-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cd205-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cd205-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cd205-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cd205-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cd205-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cd205-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cd205-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cd205-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="cd205-253">??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-253">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-254">Office ??? Android</span><span class="sxs-lookup"><span data-stu-id="cd205-254">Office for Android</span></span></td>
    <td> <span data-ttu-id="cd205-255">- ?????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-255">- Mail Read</span></span><br><span data-ttu-id="cd205-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cd205-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cd205-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cd205-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cd205-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cd205-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cd205-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="cd205-262">??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-262">not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="cd205-263">Word</span><span class="sxs-lookup"><span data-stu-id="cd205-263">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cd205-264">?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-264">Platform</span></span></th>
    <th><span data-ttu-id="cd205-265">????? ??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-265">Extension points</span></span></th> 
    <th><span data-ttu-id="cd205-266">?????? ???????????? ????????? API</span><span class="sxs-lookup"><span data-stu-id="cd205-266">API requirement sets</span></span></th> 
    <th><span data-ttu-id="cd205-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></span><span class="sxs-lookup"><span data-stu-id="cd205-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-268">Office Online</span><span class="sxs-lookup"><span data-stu-id="cd205-268">Office Online</span></span></td>
    <td> <span data-ttu-id="cd205-269">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-269">- Taskpane</span></span><br><span data-ttu-id="cd205-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cd205-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cd205-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cd205-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cd205-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-275">-BindingEvents</span></span><br><span data-ttu-id="cd205-276">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cd205-276">customXmlParts</span></span><br><span data-ttu-id="cd205-277">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-277">
         -MatrixBindings</span></span><br><span data-ttu-id="cd205-278">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-278">
         -MatrixCoercion</span></span><br><span data-ttu-id="cd205-279">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-279">
         -TableBindings</span></span><br><span data-ttu-id="cd205-280">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-280">
         -TableCoercion</span></span><br><span data-ttu-id="cd205-281">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-281">
         -TextBindings</span></span><br><span data-ttu-id="cd205-282">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-282">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-283">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cd205-283">
         -TextFile</span></span><br><span data-ttu-id="cd205-284">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-284">
         -ImageCoercion</span></span><br><span data-ttu-id="cd205-285">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-285">
         - Settings</span></span><br><span data-ttu-id="cd205-286">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-286">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-287">Office 2013 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="cd205-288">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-288">- Taskpane</span></span></td>
    <td> <span data-ttu-id="cd205-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cd205-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-290">-BindingEvents</span></span><br><span data-ttu-id="cd205-291">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-291">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-292">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="cd205-292">
         -CustomXmlPart</span></span><br><span data-ttu-id="cd205-293">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-293">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-294">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-294">
         - File</span></span><br><span data-ttu-id="cd205-295">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-295">
         -HtmlCoercion</span></span><br><span data-ttu-id="cd205-296">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-296">
         -ImageCoercion</span></span><br><span data-ttu-id="cd205-297">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-297">
         -OoxmlCoercion</span></span><br><span data-ttu-id="cd205-298">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-298">
         -TableBindings</span></span><br><span data-ttu-id="cd205-299">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-299">
         -TableCoercion</span></span><br><span data-ttu-id="cd205-300">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-300">
         -TextBindings</span></span><br><span data-ttu-id="cd205-301">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cd205-301">
         -TextFile</span></span><br><span data-ttu-id="cd205-302">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-302">
         - Settings</span></span><br><span data-ttu-id="cd205-303">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-303">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-304">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-304">
         -MatrixCoercion</span></span><br><span data-ttu-id="cd205-305">
         - ???????? ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-305">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-306">Office 2016 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-306">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="cd205-307">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-307">- Taskpane</span></span><br><span data-ttu-id="cd205-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cd205-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cd205-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cd205-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cd205-313">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-313">-BindingEvents</span></span><br><span data-ttu-id="cd205-314">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-314">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-315">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="cd205-315">
         -CustomXmlPart</span></span><br><span data-ttu-id="cd205-316">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-316">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-317">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-317">
         - File</span></span><br><span data-ttu-id="cd205-318">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-318">
         -HtmlCoercion</span></span><br><span data-ttu-id="cd205-319">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-319">
         -ImageCoercion</span></span><br><span data-ttu-id="cd205-320">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-320">
         -OoxmlCoercion</span></span><br><span data-ttu-id="cd205-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-321">
         -TableBindings</span></span><br><span data-ttu-id="cd205-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-322">
         -TableCoercion</span></span><br><span data-ttu-id="cd205-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-323">
         -TextBindings</span></span><br><span data-ttu-id="cd205-324">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cd205-324">
         -TextFile</span></span><br><span data-ttu-id="cd205-325">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-325">
         - Settings</span></span><br><span data-ttu-id="cd205-326">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-326">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-327">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-327">
         -MatrixCoercion</span></span><br><span data-ttu-id="cd205-328">
         - ???????? ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-328">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-329">Office ??? iOS</span><span class="sxs-lookup"><span data-stu-id="cd205-329">Office for iOS</span></span></td>
    <td> <span data-ttu-id="cd205-330">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-330">- Taskpane</span></span></td>
    <td> <span data-ttu-id="cd205-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cd205-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cd205-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cd205-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="cd205-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="cd205-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-335">-BindingEvents</span></span><br><span data-ttu-id="cd205-336">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-336">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-337">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="cd205-337">
         -CustomXmlPart</span></span><br><span data-ttu-id="cd205-338">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-338">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-339">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-339">
         - File</span></span><br><span data-ttu-id="cd205-340">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-340">
         -HtmlCoercion</span></span><br><span data-ttu-id="cd205-341">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-341">
         -ImageCoercion</span></span><br><span data-ttu-id="cd205-342">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-342">
         -OoxmlCoercion</span></span><br><span data-ttu-id="cd205-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-343">
         -TableBindings</span></span><br><span data-ttu-id="cd205-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-344">
         -TableCoercion</span></span><br><span data-ttu-id="cd205-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-345">
         -TextBindings</span></span><br><span data-ttu-id="cd205-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cd205-346">
         -TextFile</span></span><br><span data-ttu-id="cd205-347">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-347">
         - Settings</span></span><br><span data-ttu-id="cd205-348">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-348">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-349">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-349">
         -MatrixCoercion</span></span><br><span data-ttu-id="cd205-350">
         - ???????? ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-350">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-351">Office 2016 ??? Mac</span><span class="sxs-lookup"><span data-stu-id="cd205-351">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="cd205-352">- ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-352">- Taskpane</span></span><br><span data-ttu-id="cd205-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cd205-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cd205-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cd205-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cd205-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cd205-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="cd205-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="cd205-358">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-358">-BindingEvents</span></span><br><span data-ttu-id="cd205-359">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-359">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-360">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="cd205-360">
         -CustomXmlPart</span></span><br><span data-ttu-id="cd205-361">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-361">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-362">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-362">
         - File</span></span><br><span data-ttu-id="cd205-363">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-363">
         -HtmlCoercion</span></span><br><span data-ttu-id="cd205-364">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-364">
         -ImageCoercion</span></span><br><span data-ttu-id="cd205-365">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-365">
         -OoxmlCoercion</span></span><br><span data-ttu-id="cd205-366">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-366">
         -TableBindings</span></span><br><span data-ttu-id="cd205-367">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-367">
         -TableCoercion</span></span><br><span data-ttu-id="cd205-368">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cd205-368">
         -TextBindings</span></span><br><span data-ttu-id="cd205-369">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cd205-369">
         -TextFile</span></span><br><span data-ttu-id="cd205-370">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-370">
         - Settings</span></span><br><span data-ttu-id="cd205-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-371">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-372">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-372">
         -MatrixCoercion</span></span><br><span data-ttu-id="cd205-373">
         - ???????? ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-373">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="cd205-374">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="cd205-374">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cd205-375">?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-375">Platform</span></span></th>
    <th><span data-ttu-id="cd205-376">????? ??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-376">Extension points</span></span></th> 
    <th><span data-ttu-id="cd205-377">?????? ???????????? ????????? API</span><span class="sxs-lookup"><span data-stu-id="cd205-377">API requirement sets</span></span></th> 
    <th><span data-ttu-id="cd205-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></span><span class="sxs-lookup"><span data-stu-id="cd205-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-379">Office Online</span><span class="sxs-lookup"><span data-stu-id="cd205-379">Office Online</span></span></td>
    <td> <span data-ttu-id="cd205-380">- ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-380">- Content</span></span><br><span data-ttu-id="cd205-381">
         - ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-381">
         - Taskpane</span></span><br><span data-ttu-id="cd205-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cd205-384">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cd205-384">-ActiveView</span></span><br><span data-ttu-id="cd205-385">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-385">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-386">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-386">
         - File</span></span><br><span data-ttu-id="cd205-387">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-387">
         - Selection</span></span><br><span data-ttu-id="cd205-388">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-388">
         - Settings</span></span><br><span data-ttu-id="cd205-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-389">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-390">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-390">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-391">Office 2013 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="cd205-392">- ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-392">- Content</span></span><br><span data-ttu-id="cd205-393">
         - ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-393">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="cd205-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="cd205-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="cd205-395">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cd205-395">-ActiveView</span></span><br><span data-ttu-id="cd205-396">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-396">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-397">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-398">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-398">
         - File</span></span><br><span data-ttu-id="cd205-399">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-399">
         - Selection</span></span><br><span data-ttu-id="cd205-400">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-400">
         - Settings</span></span><br><span data-ttu-id="cd205-401">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-401">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-402">Office 2016 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-402">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="cd205-403">- ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-403">- Content</span></span><br><span data-ttu-id="cd205-404">
         - ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-404">
         - Taskpane</span></span><br><span data-ttu-id="cd205-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cd205-407">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cd205-407">-ActiveView</span></span><br><span data-ttu-id="cd205-408">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-408">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-409">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-409">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-410">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-410">
         - File</span></span><br><span data-ttu-id="cd205-411">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-411">
         - Selection</span></span><br><span data-ttu-id="cd205-412">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-412">
         - Settings</span></span><br><span data-ttu-id="cd205-413">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-413">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-414">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-414">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-415">Office ??? iOS</span><span class="sxs-lookup"><span data-stu-id="cd205-415">Office for iOS</span></span></td>
    <td> <span data-ttu-id="cd205-416">- ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-416">- Content</span></span><br><span data-ttu-id="cd205-417">
         - ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-417">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="cd205-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="cd205-419">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cd205-419">-ActiveView</span></span><br><span data-ttu-id="cd205-420">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-420">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-421">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-421">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-422">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-422">
         - File</span></span><br><span data-ttu-id="cd205-423">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-423">
         - Selection</span></span><br><span data-ttu-id="cd205-424">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-424">
         - Settings</span></span><br><span data-ttu-id="cd205-425">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-425">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-426">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-426">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-427">Office 2016 ??? Mac</span><span class="sxs-lookup"><span data-stu-id="cd205-427">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="cd205-428">- ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-428">- Content</span></span><br><span data-ttu-id="cd205-429">
         - ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-429">
         - Taskpane</span></span><br><span data-ttu-id="cd205-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cd205-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cd205-432">-ActiveView</span></span><br><span data-ttu-id="cd205-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cd205-433">
         -CompressedFile</span></span><br><span data-ttu-id="cd205-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-434">
         -DocumentEvents</span></span><br><span data-ttu-id="cd205-435">
         - ????</span><span class="sxs-lookup"><span data-stu-id="cd205-435">
         - File</span></span><br><span data-ttu-id="cd205-436">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-436">
         - Selection</span></span><br><span data-ttu-id="cd205-437">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-437">
         - Settings</span></span><br><span data-ttu-id="cd205-438">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-438">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-439">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-439">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="cd205-440">OneNote</span><span class="sxs-lookup"><span data-stu-id="cd205-440">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cd205-441">?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-441">Platform</span></span></th>
    <th><span data-ttu-id="cd205-442">????? ??????????</span><span class="sxs-lookup"><span data-stu-id="cd205-442">Extension points</span></span></th> 
    <th><span data-ttu-id="cd205-443">?????? ???????????? ????????? API</span><span class="sxs-lookup"><span data-stu-id="cd205-443">API requirement sets</span></span></th> 
    <th><span data-ttu-id="cd205-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></span><span class="sxs-lookup"><span data-stu-id="cd205-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-445">Office Online</span><span class="sxs-lookup"><span data-stu-id="cd205-445">Office Online</span></span></td>
    <td> <span data-ttu-id="cd205-446">- ???????</span><span class="sxs-lookup"><span data-stu-id="cd205-446">- Content</span></span><br><span data-ttu-id="cd205-447">
         - ??????? ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-447">
         - Taskpane</span></span><br><span data-ttu-id="cd205-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></span><span class="sxs-lookup"><span data-stu-id="cd205-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cd205-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="cd205-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cd205-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cd205-451">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cd205-451">-DocumentEvents</span></span><br><span data-ttu-id="cd205-452">
         - ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-452">
         - Settings</span></span><br><span data-ttu-id="cd205-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-453">
         -TextCoercion</span></span><br><span data-ttu-id="cd205-454">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-454">
         -HtmlCoercion</span></span><br><span data-ttu-id="cd205-455">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cd205-455">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-456">Office 2013 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-456">Office 2013 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td><span data-ttu-id="cd205-457">Office 2016 ??? Windows</span><span class="sxs-lookup"><span data-stu-id="cd205-457">Office 2016 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-458">Office ??? iOS</span><span class="sxs-lookup"><span data-stu-id="cd205-458">Office for iOS</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td><span data-ttu-id="cd205-459">Office 2016 ??? Mac</span><span class="sxs-lookup"><span data-stu-id="cd205-459">Office 2016 for Mac</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

<span data-ttu-id="cd205-460">\* = ????????? ????? ????????.</span><span class="sxs-lookup"><span data-stu-id="cd205-460">\* = We're working on it.</span></span> 

## <a name="see-also"></a><span data-ttu-id="cd205-461">??. ?????</span><span class="sxs-lookup"><span data-stu-id="cd205-461">See also</span></span>

- [<span data-ttu-id="cd205-462">????? ????????? ????????? Office</span><span class="sxs-lookup"><span data-stu-id="cd205-462">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="cd205-463">??????????? ?????? ???????????? ????????? API</span><span class="sxs-lookup"><span data-stu-id="cd205-463">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="cd205-464">?????? ???????????? ????????? ??? ?????? ?????????</span><span class="sxs-lookup"><span data-stu-id="cd205-464">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="cd205-465">??????? ?? API JavaScript ??? Office</span><span class="sxs-lookup"><span data-stu-id="cd205-465">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

