---
title: Работа с датами с помощью API JavaScript Excel
description: Используйте подключаемый Moment-MSDate с API JavaScript Excel для работы с датами.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d3f59e5daad042541bd933fb4e644d40f27a6e5e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652933"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a><span data-ttu-id="38084-103">Работа с датами с помощью API JavaScript Excel и Moment-MSDate плагина</span><span class="sxs-lookup"><span data-stu-id="38084-103">Work with dates using the Excel JavaScript API and the Moment-MSDate plug-in</span></span>

<span data-ttu-id="38084-104">В этой статье приводится пример кода, который показывает, как работать с датами с помощью API JavaScript Excel и [плагина Moment-MSDate.](https://www.npmjs.com/package/moment-msdate)</span><span class="sxs-lookup"><span data-stu-id="38084-104">This article provides code samples that show how to work with dates using the Excel JavaScript API and the [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate).</span></span> <span data-ttu-id="38084-105">Полный список свойств и методов, поддерживаемых объектом, см. `Range` в класс [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="38084-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a><span data-ttu-id="38084-106">Используйте подключаемый Moment-MSDate для работы с датами</span><span class="sxs-lookup"><span data-stu-id="38084-106">Use the Moment-MSDate plug-in to work with dates</span></span>

<span data-ttu-id="38084-107">[Библиотека JavaScript Moment](https://momentjs.com/) предоставляет удобный способ использования дат и меток времени.</span><span class="sxs-lookup"><span data-stu-id="38084-107">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="38084-108">[Подключаемый модуль Moment-MSDate](https://www.npmjs.com/package/moment-msdate) преобразует формат моментов времени в предпочитаемый для Excel.</span><span class="sxs-lookup"><span data-stu-id="38084-108">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="38084-109">Это тот же формат, который возвращает [функция ТДАТА](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46).</span><span class="sxs-lookup"><span data-stu-id="38084-109">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="38084-110">В следующем коде показано, как установить диапазон **на уровне B4** до момента.</span><span class="sxs-lookup"><span data-stu-id="38084-110">The following code shows how to set the range at **B4** to a moment's timestamp.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="38084-111">В следующем примере кода демонстрируется аналогичная техника, чтобы вернуть дату из ячейки и преобразовать ее в другой `Moment` формат.</span><span class="sxs-lookup"><span data-stu-id="38084-111">The following code sample demonstrates a similar technique to get the date back out of the cell and convert it to a `Moment` or other format.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="38084-112">Надстройка должна отформатирование диапазонов для отображения дат в более понятной для человека форме.</span><span class="sxs-lookup"><span data-stu-id="38084-112">Your add-in has to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="38084-113">Например, `"[$-409]m/d/yy h:mm AM/PM;@"` отображает "12/3/18 3:57 PM".</span><span class="sxs-lookup"><span data-stu-id="38084-113">For example, `"[$-409]m/d/yy h:mm AM/PM;@"` displays "12/3/18 3:57 PM".</span></span> <span data-ttu-id="38084-114">Дополнительные сведения о форматах дат и номеров времени см. в статье "Рекомендации по датам и форматам времени" в руководстве По обзору для настройки статьи [формата](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) номеров.</span><span class="sxs-lookup"><span data-stu-id="38084-114">For more information about date and time number formats, see "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>


## <a name="see-also"></a><span data-ttu-id="38084-115">См. также</span><span class="sxs-lookup"><span data-stu-id="38084-115">See also</span></span>

- [<span data-ttu-id="38084-116">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="38084-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="38084-117">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="38084-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="38084-118">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="38084-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
