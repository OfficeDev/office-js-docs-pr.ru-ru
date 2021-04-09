---
title: Получите диапазон с помощью API JavaScript Excel
description: Узнайте, как получить диапазон с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652927"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="eb46f-103">Получите диапазон с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="eb46f-103">Get a range using the Excel JavaScript API</span></span>

<span data-ttu-id="eb46f-104">В этой статье приводится ряд примеров получения диапазона в листах с помощью API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="eb46f-104">This article provides examples that show different ways to get a range within a worksheet using the Excel JavaScript API.</span></span> <span data-ttu-id="eb46f-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="eb46f-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a><span data-ttu-id="eb46f-106">Получение диапазона по адресу</span><span class="sxs-lookup"><span data-stu-id="eb46f-106">Get range by address</span></span>

<span data-ttu-id="eb46f-107">Следующий пример кода получает диапазон с адресом **B2:C5** из таблицы с именем **Sample,** загружает ее свойство и пишет сообщение на `address` консоль.</span><span class="sxs-lookup"><span data-stu-id="eb46f-107">The following code sample gets the range with address **B2:C5** from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a><span data-ttu-id="eb46f-108">Получение диапазона по имени</span><span class="sxs-lookup"><span data-stu-id="eb46f-108">Get range by name</span></span>

<span data-ttu-id="eb46f-109">Следующий пример кода получает диапазон, названный из таблицы с именем Sample, загружает его свойство и пишет `MyRange` сообщение на  `address` консоль.</span><span class="sxs-lookup"><span data-stu-id="eb46f-109">The following code sample gets the range named `MyRange` from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a><span data-ttu-id="eb46f-110">Получение используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="eb46f-110">Get used range</span></span>

<span data-ttu-id="eb46f-111">В следующем примере кода используется диапазон от таблицы с именем **Sample,** загружается его свойство и записывает сообщение `address` на консоль.</span><span class="sxs-lookup"><span data-stu-id="eb46f-111">The following code sample gets the used range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span> <span data-ttu-id="eb46f-112">Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки листа, которые содержат значение или форматирование.</span><span class="sxs-lookup"><span data-stu-id="eb46f-112">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="eb46f-113">Если весь лист пустой, метод возвращает диапазон, состоящий только из `getUsedRange()` верхнего левого элемента.</span><span class="sxs-lookup"><span data-stu-id="eb46f-113">If the entire worksheet is blank, the `getUsedRange()` method returns a range that consists of only the top-left cell.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a><span data-ttu-id="eb46f-114">Получение всего диапазона</span><span class="sxs-lookup"><span data-stu-id="eb46f-114">Get entire range</span></span>

<span data-ttu-id="eb46f-115">Следующий пример кода получает весь диапазон таблицы из таблицы с именем **Sample,** загружает ее свойство и пишет сообщение `address` на консоль.</span><span class="sxs-lookup"><span data-stu-id="eb46f-115">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="eb46f-116">См. также</span><span class="sxs-lookup"><span data-stu-id="eb46f-116">See also</span></span>

- [<span data-ttu-id="eb46f-117">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="eb46f-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="eb46f-118">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="eb46f-118">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="eb46f-119">Вставьте диапазон с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="eb46f-119">Insert a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-insert.md)
