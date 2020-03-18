---
title: Обзор API JavaScript для Excel
description: Узнайте больше об Excel JavaScript API
ms.date: 02/19/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d4721c349de7b93ec34ce3c35fbcee25438564db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719924"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="ed6a3-103">Обзор API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ed6a3-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="ed6a3-104">Надстройка Excel взаимодействует с объектами в Excel с помощью API JavaScript для Office, включающего две объектных модели JavaScript:</span><span class="sxs-lookup"><span data-stu-id="ed6a3-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="ed6a3-105">**API JavaScript для Excel**. Появившийся в Office 2016 [API JavaScript для Excel](/javascript/api/excel) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам.</span><span class="sxs-lookup"><span data-stu-id="ed6a3-105">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="ed6a3-106">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="ed6a3-106">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="ed6a3-107">В этом разделе рассматривается API JavaScript для Excel, используемый для разработки большинства функций в надстройках и предназначенный для Excel в Интернете, Excel 2016 или более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="ed6a3-107">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="ed6a3-108">Сведения об общем API см. в статье [Общая объектная модель API JavaScript](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="ed6a3-108">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="ed6a3-109">Сведения о концепциях, связанных с программированием</span><span class="sxs-lookup"><span data-stu-id="ed6a3-109">Learn programming concepts</span></span>

<span data-ttu-id="ed6a3-110">Сведения о важных концепциях программирования см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="ed6a3-110">See the following articles for information about important programming concepts:</span></span>
 
- [<span data-ttu-id="ed6a3-111">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ed6a3-111">Fundamental programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-core-concepts.md)

- [<span data-ttu-id="ed6a3-112">Дополнительные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ed6a3-112">Advanced programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-advanced-concepts.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="ed6a3-113">Сведения о возможностях API</span><span class="sxs-lookup"><span data-stu-id="ed6a3-113">Learn about API capabilities</span></span>

<span data-ttu-id="ed6a3-114">Используйте другие статьи этого раздела, чтобы узнать о работе с [событиями](../../excel/excel-add-ins-events.md), [диаграммами](../../excel/excel-add-ins-charts.md), [диапазонами](../../excel/excel-add-ins-ranges.md), [таблицами](../../excel/excel-add-ins-tables.md), [листами](../../excel/excel-add-ins-worksheets.md) и т. д.</span><span class="sxs-lookup"><span data-stu-id="ed6a3-114">Use other articles in this section of the documentation to learn about working with [events](../../excel/excel-add-ins-events.md), [charts](../../excel/excel-add-ins-charts.md), [ranges](../../excel/excel-add-ins-ranges.md), [tables](../../excel/excel-add-ins-tables.md), [worksheets](../../excel/excel-add-ins-worksheets.md), and more.</span></span> <span data-ttu-id="ed6a3-115">Кроме того, в этом разделе содержится руководство по концепциям API JavaScript для Excel, таким как [совместное редактирование в надстройках Excel](../../excel/co-authoring-in-excel-add-ins.md), [проверка данных](../../excel/excel-add-ins-data-validation.md), [обработка ошибок](../../excel/excel-add-ins-error-handling.md) и [оптимизация производительности](../../excel/performance.md).</span><span class="sxs-lookup"><span data-stu-id="ed6a3-115">Also in this section, you'll find guidance about Excel JavaScript API concepts such as [coauthoring in Excel add-ins](../../excel/co-authoring-in-excel-add-ins.md), [data validation](../../excel/excel-add-ins-data-validation.md), [error handling](../../excel/excel-add-ins-error-handling.md), and [performance optimization](../../excel/performance.md).</span></span> <span data-ttu-id="ed6a3-116">Полный список доступных статей см. в оглавлении.</span><span class="sxs-lookup"><span data-stu-id="ed6a3-116">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="ed6a3-117">Чтобы получить практический опыт доступа к объектам в Excel с помощью API JavaScript для Excel, выполните инструкции из [руководства по надстройкам Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="ed6a3-117">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span> 

<span data-ttu-id="ed6a3-118">Дополнительные сведения об объектной модели API JavaScript для Excel см. в [справочной документации по API JavaScript для Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="ed6a3-118">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="ed6a3-119">Опробуйте примеры кода в Script Lab</span><span class="sxs-lookup"><span data-stu-id="ed6a3-119">Try out code samples in Script Lab</span></span>

<span data-ttu-id="ed6a3-120">Используйте [Script Lab](../../overview/explore-with-script-lab.md), чтобы быстро начать работу с коллекцией встроенных примеров, демонстрирующих выполнение задач с помощью API.</span><span class="sxs-lookup"><span data-stu-id="ed6a3-120">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="ed6a3-121">Вы можете выполнять примеры в Script Lab, чтобы сразу увидеть результат в области задач или листе, изучать примеры, чтобы понять принципы действия API, и даже использовать примеры для создания собственных надстроек.</span><span class="sxs-lookup"><span data-stu-id="ed6a3-121">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="ed6a3-122">См. также</span><span class="sxs-lookup"><span data-stu-id="ed6a3-122">See also</span></span>

- [<span data-ttu-id="ed6a3-123">Документация по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="ed6a3-123">Excel add-ins documentation</span></span>](../../excel/index.md)
- [<span data-ttu-id="ed6a3-124">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="ed6a3-124">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
- [<span data-ttu-id="ed6a3-125">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ed6a3-125">Excel JavaScript API reference</span></span>](/javascript/api/excel)
- [<span data-ttu-id="ed6a3-126">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ed6a3-126">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="ed6a3-127">Открытые спецификации API</span><span class="sxs-lookup"><span data-stu-id="ed6a3-127">API open specifications</span></span>](../openspec/openspec.md)
