---
title: Обзор API JavaScript для Excel
description: ''
ms.date: 02/19/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aa0e18025e539f747851f5dc9f5a25197761c5c8
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163621"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="25e62-102">Обзор API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="25e62-102">Excel JavaScript API overview</span></span>

<span data-ttu-id="25e62-103">Надстройка Excel взаимодействует с объектами в Excel с помощью API JavaScript для Office, включающего две объектных модели JavaScript:</span><span class="sxs-lookup"><span data-stu-id="25e62-103">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="25e62-104">**API JavaScript для Excel**. Появившийся в Office 2016 [API JavaScript для Excel](/javascript/api/excel) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам.</span><span class="sxs-lookup"><span data-stu-id="25e62-104">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="25e62-105">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="25e62-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="25e62-106">В этом разделе рассматривается API JavaScript для Excel, используемый для разработки большинства функций в надстройках и предназначенный для Excel в Интернете, Excel 2016 или более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="25e62-106">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="25e62-107">Сведения об общем API см. в статье [Общая объектная модель API JavaScript](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="25e62-107">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="25e62-108">Сведения о концепциях, связанных с программированием</span><span class="sxs-lookup"><span data-stu-id="25e62-108">Learn programming concepts</span></span>

<span data-ttu-id="25e62-109">Сведения о важных концепциях программирования см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="25e62-109">See the following articles for information about important programming concepts:</span></span>
 
- [<span data-ttu-id="25e62-110">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="25e62-110">Fundamental programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-core-concepts.md)

- [<span data-ttu-id="25e62-111">Дополнительные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="25e62-111">Advanced programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-advanced-concepts.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="25e62-112">Сведения о возможностях API</span><span class="sxs-lookup"><span data-stu-id="25e62-112">Learn about API capabilities</span></span>

<span data-ttu-id="25e62-113">Используйте другие статьи этого раздела, чтобы узнать о работе с [событиями](../../excel/excel-add-ins-events.md), [диаграммами](../../excel/excel-add-ins-charts.md), [диапазонами](../../excel/excel-add-ins-ranges.md), [таблицами](../../excel/excel-add-ins-tables.md), [листами](../../excel/excel-add-ins-worksheets.md) и т. д.</span><span class="sxs-lookup"><span data-stu-id="25e62-113">Use other articles in this section of the documentation to learn about working with [events](../../excel/excel-add-ins-events.md), [charts](../../excel/excel-add-ins-charts.md), [ranges](../../excel/excel-add-ins-ranges.md), [tables](../../excel/excel-add-ins-tables.md), [worksheets](../../excel/excel-add-ins-worksheets.md), and more.</span></span> <span data-ttu-id="25e62-114">Кроме того, в этом разделе содержится руководство по концепциям API JavaScript для Excel, таким как [совместное редактирование в надстройках Excel](../../excel/co-authoring-in-excel-add-ins.md), [проверка данных](../../excel/excel-add-ins-data-validation.md), [обработка ошибок](../../excel/excel-add-ins-error-handling.md) и [оптимизация производительности](../../excel/performance.md).</span><span class="sxs-lookup"><span data-stu-id="25e62-114">Also in this section, you'll find guidance about Excel JavaScript API concepts such as [coauthoring in Excel add-ins](../../excel/co-authoring-in-excel-add-ins.md), [data validation](../../excel/excel-add-ins-data-validation.md), [error handling](../../excel/excel-add-ins-error-handling.md), and [performance optimization](../../excel/performance.md).</span></span> <span data-ttu-id="25e62-115">Полный список доступных статей см. в оглавлении.</span><span class="sxs-lookup"><span data-stu-id="25e62-115">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="25e62-116">Чтобы получить практический опыт доступа к объектам в Excel с помощью API JavaScript для Excel, выполните инструкции из [руководства по надстройкам Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="25e62-116">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span> 

<span data-ttu-id="25e62-117">Дополнительные сведения об объектной модели API JavaScript для Excel см. в [справочной документации по API JavaScript для Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="25e62-117">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="25e62-118">Опробуйте примеры кода в Script Lab</span><span class="sxs-lookup"><span data-stu-id="25e62-118">Try out code samples in Script Lab</span></span>

<span data-ttu-id="25e62-119">Используйте [Script Lab](../../overview/explore-with-script-lab.md), чтобы быстро начать работу с коллекцией встроенных примеров, демонстрирующих выполнение задач с помощью API.</span><span class="sxs-lookup"><span data-stu-id="25e62-119">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="25e62-120">Вы можете выполнять примеры в Script Lab, чтобы сразу увидеть результат в области задач или листе, изучать примеры, чтобы понять принципы действия API, и даже использовать примеры для создания собственных надстроек.</span><span class="sxs-lookup"><span data-stu-id="25e62-120">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="25e62-121">См. также</span><span class="sxs-lookup"><span data-stu-id="25e62-121">See also</span></span>

- [<span data-ttu-id="25e62-122">Документация по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="25e62-122">Excel add-ins documentation</span></span>](../../excel/index.md)
- [<span data-ttu-id="25e62-123">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="25e62-123">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
- [<span data-ttu-id="25e62-124">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="25e62-124">Excel JavaScript API reference</span></span>](/javascript/api/excel)
- [<span data-ttu-id="25e62-125">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="25e62-125">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="25e62-126">Открытые спецификации API</span><span class="sxs-lookup"><span data-stu-id="25e62-126">API open specifications</span></span>](../openspec/openspec.md)
