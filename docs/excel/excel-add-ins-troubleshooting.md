---
title: Устранение неполадок надстроек Excel
description: Узнайте, как устранять ошибки разработки в надстройках Excel.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409412"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="5d50f-103">Устранение неполадок надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="5d50f-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="5d50f-104">В этой статье обсуждаются проблемы, связанные с устранением неполадок, которые являются уникальными для Excel.</span><span class="sxs-lookup"><span data-stu-id="5d50f-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="5d50f-105">Воспользуйтесь средством обратной связи в нижней части страницы, чтобы предложить другие проблемы, которые можно добавить в статью.</span><span class="sxs-lookup"><span data-stu-id="5d50f-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="5d50f-106">Ограничения API при использовании активных переключателей книги</span><span class="sxs-lookup"><span data-stu-id="5d50f-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="5d50f-107">Надстройки для Excel предназначены для работы с одной книгой за раз.</span><span class="sxs-lookup"><span data-stu-id="5d50f-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="5d50f-108">Ошибки могут возникать, если книга, отделяющая от того, где работает надстройка, получает фокус.</span><span class="sxs-lookup"><span data-stu-id="5d50f-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="5d50f-109">Это происходит только в том случае, если определенные методы находятся в процессе вызова при изменении фокуса.</span><span class="sxs-lookup"><span data-stu-id="5d50f-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="5d50f-110">Этот переключатель книги влияет на следующие API:</span><span class="sxs-lookup"><span data-stu-id="5d50f-110">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="5d50f-111">API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="5d50f-111">Excel JavaScript API</span></span> | <span data-ttu-id="5d50f-112">Выдается ошибка</span><span class="sxs-lookup"><span data-stu-id="5d50f-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="5d50f-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="5d50f-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="5d50f-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="5d50f-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="5d50f-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="5d50f-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="5d50f-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="5d50f-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="5d50f-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="5d50f-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="5d50f-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="5d50f-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="5d50f-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="5d50f-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="5d50f-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="5d50f-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="5d50f-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="5d50f-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="5d50f-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="5d50f-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="5d50f-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="5d50f-129">Это относится только к нескольким книгам Excel, открываемым в Windows или Mac.</span><span class="sxs-lookup"><span data-stu-id="5d50f-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="5d50f-130">Совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="5d50f-130">Coauthoring</span></span>

<span data-ttu-id="5d50f-131">Используйте совместное [Редактирование в](co-authoring-in-excel-add-ins.md) надстройках Excel для шаблонов, используемых с событиями в среде совместной работы.</span><span class="sxs-lookup"><span data-stu-id="5d50f-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="5d50f-132">В этой статье также обсуждаются потенциальные конфликты объединения при использовании определенных API, например [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="5d50f-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="5d50f-133">См. также</span><span class="sxs-lookup"><span data-stu-id="5d50f-133">See also</span></span>

- [<span data-ttu-id="5d50f-134">Устранение ошибок разработки надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5d50f-134">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="5d50f-135">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="5d50f-135">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
