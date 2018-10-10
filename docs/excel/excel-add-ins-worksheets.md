---
title: Работа с листами с использованием API JavaScript для Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 9ceb2187cdd7f503fb39171e420adabcc2f13041
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459135"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="9edec-102">Работа с листами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="9edec-102">Work with Worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="9edec-103">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для листов с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="9edec-103">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span> <span data-ttu-id="9edec-104">Полный список свойств и методов, поддерживаемых объектами **Worksheet** и **WorksheetCollection**, см. в статьях [Объект Worksheet (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) и [Объект WorksheetCollection (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="9edec-104">For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) and [WorksheetCollection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection?view=office-js).</span></span>

> [!NOTE]
> <span data-ttu-id="9edec-105">Сведения в этой статье применимы только к обычным листам, а не к листам диаграмм или макросов.</span><span class="sxs-lookup"><span data-stu-id="9edec-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="9edec-106">Получение листов</span><span class="sxs-lookup"><span data-stu-id="9edec-106">Get worksheets</span></span>

<span data-ttu-id="9edec-107">В примере ниже показано, как возвратить коллекцию листов, загрузить свойство **name** каждого листа и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length > 1) {
                console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
            } else {
                console.log(`There is one worksheet in the workbook:`);
            }
            for (var i in sheets.items) {
                console.log(sheets.items[i].name);
            }
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="9edec-108">Свойство **id** листа уникальным образом идентифицирует лист в конкретной книге, и его значение не изменяется даже при переименовании или перемещении листа.</span><span class="sxs-lookup"><span data-stu-id="9edec-108">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved.</span></span> <span data-ttu-id="9edec-109">При удалении листа из книги в Excel для Mac **идентификатор** удаленного листа можно назначить новому листу (созданному после удаления).</span><span class="sxs-lookup"><span data-stu-id="9edec-109">When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="9edec-110">Получение активного листа</span><span class="sxs-lookup"><span data-stu-id="9edec-110">Get the active worksheet</span></span>

<span data-ttu-id="9edec-111">В примере кода ниже показано, как получить активный лист, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    
    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-the-active-worksheet"></a><span data-ttu-id="9edec-112">Задание активного листа</span><span class="sxs-lookup"><span data-stu-id="9edec-112">Set the active worksheet</span></span>

<span data-ttu-id="9edec-113">В примере кода ниже показано, как задать лист **Sample** (Пример) в качестве активного, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-113">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="9edec-114">Если нет листа с таким именем, метод **activate()** создаст ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="9edec-114">If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.activate();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="9edec-115">Ссылка на листы по их относительным положениям</span><span class="sxs-lookup"><span data-stu-id="9edec-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="9edec-116">В примерах ниже показано, как ссылаться на лист по его относительному положению.</span><span class="sxs-lookup"><span data-stu-id="9edec-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="9edec-117">Получение первого листа</span><span class="sxs-lookup"><span data-stu-id="9edec-117">Get the first worksheet</span></span>

<span data-ttu-id="9edec-118">В примере кода ниже показано, как получить первый лист в книге, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var firstSheet = context.workbook.worksheets.getFirst();
    firstSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the first worksheet is "${firstSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-last-worksheet"></a><span data-ttu-id="9edec-119">Получение последнего листа</span><span class="sxs-lookup"><span data-stu-id="9edec-119">Get the last worksheet</span></span>

<span data-ttu-id="9edec-120">В примере кода ниже показано, как получить последний лист в книге, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var lastSheet = context.workbook.worksheets.getLast();
    lastSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the last worksheet is "${lastSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-next-worksheet"></a><span data-ttu-id="9edec-121">Получение следующего листа</span><span class="sxs-lookup"><span data-stu-id="9edec-121">Get the next worksheet</span></span>

<span data-ttu-id="9edec-122">В примере кода ниже показано, как получить лист, следующий за активным листом в книге, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-122">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="9edec-123">Если нет листа после активного листа, метод **getNext()** создаст ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="9edec-123">If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

```js
 Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var nextSheet = currentSheet.getNext();
    nextSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="9edec-124">Получение предыдущего листа</span><span class="sxs-lookup"><span data-stu-id="9edec-124">Get the previous worksheet</span></span>

<span data-ttu-id="9edec-125">В примере кода ниже показано, как получить лист, предшествующий активному листу в книге, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-125">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="9edec-126">Если нет листа перед активным листом, метод **getPrevious()** создаст ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="9edec-126">If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var previousSheet = currentSheet.getPrevious();
    previousSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="add-a-worksheet"></a><span data-ttu-id="9edec-127">Добавление листа</span><span class="sxs-lookup"><span data-stu-id="9edec-127">Add a worksheet</span></span>

<span data-ttu-id="9edec-p106">В примере кода ниже показано, как добавить лист **Sample** (Пример) в рабочую книгу, загрузить его свойства **name** и **position** и записать сообщение в консоль. Новый лист будет следовать за всеми остальными.</span><span class="sxs-lookup"><span data-stu-id="9edec-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;

    var sheet = sheets.add("Sample");
    sheet.load("name, position");
    
    return context.sync()
        .then(function () {
            console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
        });
}).catch(errorHandlerFunction);
```

## <a name="delete-a-worksheet"></a><span data-ttu-id="9edec-130">Удаление листа</span><span class="sxs-lookup"><span data-stu-id="9edec-130">Delete a worksheet</span></span>

<span data-ttu-id="9edec-131">В примере кода ниже показано, как удалить последний лист в книге (если это не единственный лист в книге) и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length === 1) {
                console.log("Unable to delete the only worksheet in the workbook");
            } else {
                var lastSheet = sheets.items[sheets.items.length - 1];

                console.log(`Deleting worksheet named "${lastSheet.name}"`);
                lastSheet.delete();

                return context.sync();
            };
        });
}).catch(errorHandlerFunction);
```

## <a name="rename-a-worksheet"></a><span data-ttu-id="9edec-132">Переименование листа</span><span class="sxs-lookup"><span data-stu-id="9edec-132">Rename a worksheet</span></span>

<span data-ttu-id="9edec-133">В примере ниже показано, как изменить имя активного листа на **New Name** (Новое имя).</span><span class="sxs-lookup"><span data-stu-id="9edec-133">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="9edec-134">Перемещение листа</span><span class="sxs-lookup"><span data-stu-id="9edec-134">Move a worksheet</span></span>

<span data-ttu-id="9edec-135">В примере ниже показано, как переместить лист из последней позиции в книге на первую.</span><span class="sxs-lookup"><span data-stu-id="9edec-135">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items");

    return context.sync()
        .then(function () {
            var lastSheet = sheets.items[sheets.items.length - 1];
            lastSheet.position = 0;

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

## <a name="set-worksheet-visibility"></a><span data-ttu-id="9edec-136">Настройка видимости листа</span><span class="sxs-lookup"><span data-stu-id="9edec-136">Set worksheet visibility</span></span>

<span data-ttu-id="9edec-137">В примерах ниже показано, как настроить видимость листа.</span><span class="sxs-lookup"><span data-stu-id="9edec-137">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="9edec-138">Скрытие листа</span><span class="sxs-lookup"><span data-stu-id="9edec-138">Hide a worksheet</span></span>

<span data-ttu-id="9edec-139">В примере кода ниже показано, как сделать лист **Sample** (Пример) скрытым, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-139">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.hidden;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is hidden`);
        });
}).catch(errorHandlerFunction);
```

### <a name="unhide-a-worksheet"></a><span data-ttu-id="9edec-140">Отмена скрытия листа</span><span class="sxs-lookup"><span data-stu-id="9edec-140">Unhide a worksheet</span></span>

<span data-ttu-id="9edec-141">В примере кода ниже показано, как сделать лист **Sample** (Пример), загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-141">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is visible`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-a-cell-within-a-worksheet"></a><span data-ttu-id="9edec-142">Получение ячейки листа</span><span class="sxs-lookup"><span data-stu-id="9edec-142">Get a cell within a worksheet</span></span>

<span data-ttu-id="9edec-143">В примере кода ниже показано, как получить ячейку, расположенную в строке 2 и столбце 5 листа **Sample** (Пример), загрузить ее свойства **address** и **values** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="9edec-143">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="9edec-144">Значения, передаваемые в метод **getCell(row: число, column: число)**, представляют собой индексируемые с нуля номера строк и столбцов получаемой ячейки.</span><span class="sxs-lookup"><span data-stu-id="9edec-144">The values that are passed into the **getCell(row: number, column:number)** method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var cell = sheet.getCell(1, 4);
    cell.load("address, values");
    
    return context.sync()
        .then(function() {
            console.log(`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`);
        })
}).catch(errorHandlerFunction);
```

## <a name="get-a-range-within-a-worksheet"></a><span data-ttu-id="9edec-145">Получение диапазона на листе</span><span class="sxs-lookup"><span data-stu-id="9edec-145">Get a range within a worksheet</span></span>

<span data-ttu-id="9edec-146">Примеры, в которых показано, как получить диапазон на листе, см. в статье [Работа с диапазонами с использованием API JavaScript для Excel](excel-add-ins-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="9edec-146">For examples that show how to get a range within a worksheet, see [Work with Ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="9edec-147">См. также</span><span class="sxs-lookup"><span data-stu-id="9edec-147">See also</span></span>

- [<span data-ttu-id="9edec-148">Основные принципы программирования с использованием интерфейса API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="9edec-148">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

