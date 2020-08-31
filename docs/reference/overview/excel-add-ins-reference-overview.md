---
title: Обзор API JavaScript для Excel
description: Узнайте больше об Excel JavaScript API
ms.date: 07/28/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e589bd7ce814211759cc731d828e9c180339ea1f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293662"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="e1097-103">Обзор API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e1097-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="e1097-104">Надстройка Excel взаимодействует с объектами в Excel с помощью API JavaScript для Office, включающего две объектных модели JavaScript:</span><span class="sxs-lookup"><span data-stu-id="e1097-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="e1097-105">**API JavaScript для Excel**. Это [API-интерфейсы для определенных приложений](../../develop/application-specific-api-model.md) в Excel.</span><span class="sxs-lookup"><span data-stu-id="e1097-105">**Excel JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Excel.</span></span> <span data-ttu-id="e1097-106">Появившийся в Office 2016 [API JavaScript для Excel](/javascript/api/excel) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам.</span><span class="sxs-lookup"><span data-stu-id="e1097-106">Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="e1097-107">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="e1097-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="e1097-108">В этом разделе рассматривается API JavaScript для Excel, используемый для разработки большинства функций в надстройках и предназначенный для Excel в Интернете, Excel 2016 или более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="e1097-108">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="e1097-109">Сведения об общем API см. в статье [Общая объектная модель API JavaScript](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="e1097-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="e1097-110">Сведения о концепциях, связанных с программированием</span><span class="sxs-lookup"><span data-stu-id="e1097-110">Learn programming concepts</span></span>

<span data-ttu-id="e1097-111">Сведения о важных концепциях программирования см. в статье [Основные концепции программирования с помощью API JavaScript для Excel](../../excel/excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="e1097-111">See [Fundamental programming concepts with the Excel JavaScript API](../../excel/excel-add-ins-core-concepts.md) for information about important programming concepts.</span></span>

<span data-ttu-id="e1097-112">Чтобы получить практический опыт доступа к объектам в Excel с помощью API JavaScript для Excel, выполните инструкции из [руководства по надстройкам Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="e1097-112">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span>

## <a name="learn-api-capabilities"></a><span data-ttu-id="e1097-113">Сведения о возможностях API</span><span class="sxs-lookup"><span data-stu-id="e1097-113">Learn API capabilities</span></span>

<span data-ttu-id="e1097-114">Каждой основной функции API Excel посвящена статья, описывающая ее возможности и соответствующую объектную модель.</span><span class="sxs-lookup"><span data-stu-id="e1097-114">Each major Excel API feature has an article exploring what that feature can do and the relevant object model.</span></span>

* [<span data-ttu-id="e1097-115">Диаграммы</span><span class="sxs-lookup"><span data-stu-id="e1097-115">Charts</span></span>](../../excel/excel-add-ins-charts.md)
* [<span data-ttu-id="e1097-116">Комментарии</span><span class="sxs-lookup"><span data-stu-id="e1097-116">Comments</span></span>](../../excel/excel-add-ins-comments.md)
* [<span data-ttu-id="e1097-117">Условное форматирование</span><span class="sxs-lookup"><span data-stu-id="e1097-117">Conditional formatting</span></span>](../../excel/excel-add-ins-conditional-formatting.md)
* [<span data-ttu-id="e1097-118">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="e1097-118">Custom functions</span></span>](../../excel/custom-functions-overview.md)
* [<span data-ttu-id="e1097-119">Проверка данных</span><span class="sxs-lookup"><span data-stu-id="e1097-119">Data validation</span></span>](../../excel/excel-add-ins-data-validation.md)
* [<span data-ttu-id="e1097-120">События</span><span class="sxs-lookup"><span data-stu-id="e1097-120">Events</span></span>](../../excel/excel-add-ins-events.md)
* [<span data-ttu-id="e1097-121">Несколько диапазонов (RangeArea)</span><span class="sxs-lookup"><span data-stu-id="e1097-121">Multiple ranges (RangeArea)</span></span>](../../excel/excel-add-ins-multiple-ranges.md)
* [<span data-ttu-id="e1097-122">Сводные таблицы</span><span class="sxs-lookup"><span data-stu-id="e1097-122">PivotTables</span></span>](../../excel/excel-add-ins-pivottables.md)
* <span data-ttu-id="e1097-123">[Диапазоны](../../excel/excel-add-ins-ranges.md) и [API расширенных диапазонов](../../excel/excel-add-ins-ranges-advanced.md)</span><span class="sxs-lookup"><span data-stu-id="e1097-123">[Ranges](../../excel/excel-add-ins-ranges.md) and [Advanced Range APIs](../../excel/excel-add-ins-ranges-advanced.md)</span></span>
* [<span data-ttu-id="e1097-124">Фигуры</span><span class="sxs-lookup"><span data-stu-id="e1097-124">Shapes</span></span>](../../excel/excel-add-ins-shapes.md)
* [<span data-ttu-id="e1097-125">Таблицы</span><span class="sxs-lookup"><span data-stu-id="e1097-125">Tables</span></span>](../../excel/excel-add-ins-tables.md)
* [<span data-ttu-id="e1097-126">Книги и API уровня приложения</span><span class="sxs-lookup"><span data-stu-id="e1097-126">Workbooks and Application-level APIs</span></span>](../../excel/excel-add-ins-workbooks.md)
* [<span data-ttu-id="e1097-127">Листы</span><span class="sxs-lookup"><span data-stu-id="e1097-127">Worksheets</span></span>](../../excel/excel-add-ins-worksheets.md)

<span data-ttu-id="e1097-128">Дополнительные сведения об объектной модели API JavaScript для Excel см. в [справочной документации по API JavaScript для Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="e1097-128">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="e1097-129">Опробуйте примеры кода в Script Lab</span><span class="sxs-lookup"><span data-stu-id="e1097-129">Try out code samples in Script Lab</span></span>

<span data-ttu-id="e1097-130">Используйте [Script Lab](../../overview/explore-with-script-lab.md), чтобы быстро начать работу с коллекцией встроенных примеров, демонстрирующих выполнение задач с помощью API.</span><span class="sxs-lookup"><span data-stu-id="e1097-130">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="e1097-131">Вы можете выполнять примеры в Script Lab, чтобы сразу увидеть результат в области задач или листе, изучать примеры, чтобы понять принципы действия API, и даже использовать примеры для создания собственных надстроек.</span><span class="sxs-lookup"><span data-stu-id="e1097-131">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="e1097-132">См. также</span><span class="sxs-lookup"><span data-stu-id="e1097-132">See also</span></span>

* [<span data-ttu-id="e1097-133">Документация по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="e1097-133">Excel add-ins documentation</span></span>](../../excel/index.yml)
* [<span data-ttu-id="e1097-134">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="e1097-134">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
* [<span data-ttu-id="e1097-135">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e1097-135">Excel JavaScript API reference</span></span>](/javascript/api/excel)
* [<span data-ttu-id="e1097-136">Доступность клиентских приложений и платформ Office для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e1097-136">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
* [<span data-ttu-id="e1097-137">Использование модели API для определенных приложений</span><span class="sxs-lookup"><span data-stu-id="e1097-137">Using the application-specific API model</span></span>](../../develop/application-specific-api-model.md)
