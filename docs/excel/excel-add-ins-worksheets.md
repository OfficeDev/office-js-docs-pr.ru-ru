---
title: Работа с листами с использованием API JavaScript для Excel
description: Примеры кода, которые показывают, как выполнять общие задачи с листами с помощью API JavaScript Excel.
ms.date: 03/24/2020
localization_priority: Normal
ms.openlocfilehash: 7ff1593ca66926de7ae3397defba7efbe97b1695
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652204"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="734e0-103">Работа с листами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="734e0-103">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="734e0-p101">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для листов с использованием API JavaScript для Excel. Полный список свойств и методов, поддерживаемых объектами `Worksheet` и `WorksheetCollection`, см. в статьях [Объект Worksheet (API JavaScript для Excel)](/javascript/api/excel/excel.worksheet) и [Объект WorksheetCollection (API JavaScript для Excel)](/javascript/api/excel/excel.worksheetcollection).</span><span class="sxs-lookup"><span data-stu-id="734e0-p101">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API. For the complete list of properties and methods that the `Worksheet` and `WorksheetCollection` objects support, see [Worksheet Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="734e0-106">Сведения в этой статье применимы только к обычным листам, а не к листам диаграмм или макросов.</span><span class="sxs-lookup"><span data-stu-id="734e0-106">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="734e0-107">Получение листов</span><span class="sxs-lookup"><span data-stu-id="734e0-107">Get worksheets</span></span>

<span data-ttu-id="734e0-108">В примере кода ниже показано, как возвратить коллекцию листов, загрузить свойство `name` каждого листа и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="734e0-108">The following code sample gets the collection of worksheets, loads the `name` property of each worksheet, and writes a message to the console.</span></span>

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
            sheets.items.forEach(function (sheet) {
              console.log(sheet.name);
            });
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="734e0-p102">Свойство `id` листа уникальным образом идентифицирует лист в конкретной книге, и его значение не изменяется даже при переименовании или перемещении листа. При удалении листа из книги в Excel для Mac `id` удаленного листа можно назначить новому листу (созданному после удаления).</span><span class="sxs-lookup"><span data-stu-id="734e0-p102">The `id` property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved. When a worksheet is deleted from a workbook in Excel on Mac, the `id` of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="734e0-111">Получение активного листа</span><span class="sxs-lookup"><span data-stu-id="734e0-111">Get the active worksheet</span></span>

<span data-ttu-id="734e0-112">В примере кода ниже показано, как получить активный лист, загрузить его свойство `name` и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="734e0-112">The following code sample gets the active worksheet, loads its `name` property, and writes a message to the console.</span></span>

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

## <a name="set-the-active-worksheet"></a><span data-ttu-id="734e0-113">Задание активного листа</span><span class="sxs-lookup"><span data-stu-id="734e0-113">Set the active worksheet</span></span>

<span data-ttu-id="734e0-p103">В примере кода ниже показано, как задать лист **Sample** (Пример) в качестве активного, загрузить его свойство `name` и записать сообщение в консоль. Если нет листа с таким именем, метод `activate()` создаст ошибку `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="734e0-p103">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its `name` property, and writes a message to the console. If there is no worksheet with that name, the `activate()` method throws an `ItemNotFound` error.</span></span>

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

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="734e0-116">Ссылка на листы по их относительным положениям</span><span class="sxs-lookup"><span data-stu-id="734e0-116">Reference worksheets by relative position</span></span>

<span data-ttu-id="734e0-117">В примерах ниже показано, как ссылаться на лист по его относительному положению.</span><span class="sxs-lookup"><span data-stu-id="734e0-117">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="734e0-118">Получение первого листа</span><span class="sxs-lookup"><span data-stu-id="734e0-118">Get the first worksheet</span></span>

<span data-ttu-id="734e0-119">В примере кода ниже показано, как получить первый лист в книге, загрузить его свойство `name` и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="734e0-119">The following code sample gets the first worksheet in the workbook, loads its `name` property, and writes a message to the console.</span></span>

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

### <a name="get-the-last-worksheet"></a><span data-ttu-id="734e0-120">Получение последнего листа</span><span class="sxs-lookup"><span data-stu-id="734e0-120">Get the last worksheet</span></span>

<span data-ttu-id="734e0-121">В примере кода ниже показано, как получить последний лист в книге, загрузить его свойство `name` и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="734e0-121">The following code sample gets the last worksheet in the workbook, loads its `name` property, and writes a message to the console.</span></span>

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

### <a name="get-the-next-worksheet"></a><span data-ttu-id="734e0-122">Получение следующего листа</span><span class="sxs-lookup"><span data-stu-id="734e0-122">Get the next worksheet</span></span>

<span data-ttu-id="734e0-p104">В примере кода ниже показано, как получить лист, следующий за активным листом, в книге, загрузить его свойство `name` и записать сообщение в консоль. Если нет листа после активного листа, метод `getNext()` создаст ошибку `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="734e0-p104">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its `name` property, and writes a message to the console. If there is no worksheet after the active worksheet, the `getNext()` method throws an `ItemNotFound` error.</span></span>

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

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="734e0-125">Получение предыдущего листа</span><span class="sxs-lookup"><span data-stu-id="734e0-125">Get the previous worksheet</span></span>

<span data-ttu-id="734e0-p105">В примере кода ниже показано, как получить лист, предшествующий активному листу, в книге, загрузить его свойство `name` и записать сообщение в консоль. Если нет листа перед активным листом, метод `getPrevious()` создаст ошибку `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="734e0-p105">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its `name` property, and writes a message to the console. If there is no worksheet before the active worksheet, the `getPrevious()` method throws an `ItemNotFound` error.</span></span>

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

## <a name="add-a-worksheet"></a><span data-ttu-id="734e0-128">Добавление листа</span><span class="sxs-lookup"><span data-stu-id="734e0-128">Add a worksheet</span></span>

<span data-ttu-id="734e0-p106">В примере кода ниже показано, как добавить лист **Sample** (Пример) в рабочую книгу, загрузить его свойства `name` и `position` и записать сообщение в консоль. Новый лист будет следовать за всеми остальными.</span><span class="sxs-lookup"><span data-stu-id="734e0-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its `name` and `position` properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

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

### <a name="copy-an-existing-worksheet"></a><span data-ttu-id="734e0-131">Копирование существующего листа</span><span class="sxs-lookup"><span data-stu-id="734e0-131">Copy an existing worksheet</span></span>

<span data-ttu-id="734e0-132">`Worksheet.copy` добавляет новый лист, являющийся копией существующего листа.</span><span class="sxs-lookup"><span data-stu-id="734e0-132">`Worksheet.copy` adds a new worksheet that is a copy of an existing worksheet.</span></span> <span data-ttu-id="734e0-133">Имя нового листа будет содержать номер в конце по аналогии с копированием листов в пользовательском интерфейсе Excel (например, **МойЛист (2)**).</span><span class="sxs-lookup"><span data-stu-id="734e0-133">The new worksheet's name will have a number appended to the end, in a manner consistent with copying a worksheet through the Excel UI (for example, **MySheet (2)**).</span></span> <span data-ttu-id="734e0-134">`Worksheet.copy` может принимать два необязательных параметра:</span><span class="sxs-lookup"><span data-stu-id="734e0-134">`Worksheet.copy` can take two parameters, both of which are optional:</span></span>

- <span data-ttu-id="734e0-135">`positionType` — перечисление [WorksheetPositionType](/javascript/api/excel/excel.worksheetpositiontype), указывающее, где в книге нужно добавить новый лист.</span><span class="sxs-lookup"><span data-stu-id="734e0-135">`positionType` - A [WorksheetPositionType](/javascript/api/excel/excel.worksheetpositiontype) enum specifying where in the workbook the new worksheet is to be added.</span></span>
- <span data-ttu-id="734e0-136">`relativeTo` — если параметру `positionType` присвоено значение `Before` или `After`, требуется указать лист, относительно которого нужно добавить новый лист (этот параметр отвечает на вопрос "До или после чего?").</span><span class="sxs-lookup"><span data-stu-id="734e0-136">`relativeTo` - If the `positionType` is `Before` or `After`, you need to specify a worksheet relative to which the new sheet is to be added (this parameter answers the question "Before or after what?").</span></span>

<span data-ttu-id="734e0-137">В примере кода ниже показано, как скопировать текущий лист и вставить новый лист непосредственно после текущего.</span><span class="sxs-lookup"><span data-stu-id="734e0-137">The following code sample copies the current worksheet and inserts the new sheet directly after the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var myWorkbook = context.workbook;
    var sampleSheet = myWorkbook.worksheets.getActiveWorksheet();
    var copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);
    return context.sync();
});
```

## <a name="delete-a-worksheet"></a><span data-ttu-id="734e0-138">Удаление листа</span><span class="sxs-lookup"><span data-stu-id="734e0-138">Delete a worksheet</span></span>

<span data-ttu-id="734e0-139">В примере кода ниже показано, как удалить последний лист в книге (если это не единственный лист в книге) и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="734e0-139">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

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

> [!NOTE]
> <span data-ttu-id="734e0-140">Лист с уровнем скрытия "[надежно скрыт](/javascript/api/excel/excel.sheetvisibility)" невозможно удалить с помощью метода `delete`.</span><span class="sxs-lookup"><span data-stu-id="734e0-140">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="734e0-141">Чтобы удалить лист, нужно сперва изменить его уровень скрытия.</span><span class="sxs-lookup"><span data-stu-id="734e0-141">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="734e0-142">Переименование листа</span><span class="sxs-lookup"><span data-stu-id="734e0-142">Rename a worksheet</span></span>

<span data-ttu-id="734e0-143">В примере ниже показано, как изменить имя активного листа на **New Name** (Новое имя).</span><span class="sxs-lookup"><span data-stu-id="734e0-143">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="734e0-144">Перемещение листа</span><span class="sxs-lookup"><span data-stu-id="734e0-144">Move a worksheet</span></span>

<span data-ttu-id="734e0-145">В примере ниже показано, как переместить лист из последней позиции в книге на первую.</span><span class="sxs-lookup"><span data-stu-id="734e0-145">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

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

## <a name="set-worksheet-visibility"></a><span data-ttu-id="734e0-146">Настройка видимости листа</span><span class="sxs-lookup"><span data-stu-id="734e0-146">Set worksheet visibility</span></span>

<span data-ttu-id="734e0-147">В примерах ниже показано, как настроить видимость листа.</span><span class="sxs-lookup"><span data-stu-id="734e0-147">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="734e0-148">Скрытие листа</span><span class="sxs-lookup"><span data-stu-id="734e0-148">Hide a worksheet</span></span>

<span data-ttu-id="734e0-149">В примере кода ниже показано, как сделать лист **Sample** (Пример) скрытым, загрузить его свойство `name` и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="734e0-149">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its `name` property, and writes a message to the console.</span></span>

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

### <a name="unhide-a-worksheet"></a><span data-ttu-id="734e0-150">Отмена скрытия листа</span><span class="sxs-lookup"><span data-stu-id="734e0-150">Unhide a worksheet</span></span>

<span data-ttu-id="734e0-151">В примере кода ниже показано, как сделать лист **Sample** (Пример) видимым, загрузить его свойство `name` и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="734e0-151">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its `name` property, and writes a message to the console.</span></span>

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

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="734e0-152">Получение одной ячейки листа</span><span class="sxs-lookup"><span data-stu-id="734e0-152">Get a single cell within a worksheet</span></span>

<span data-ttu-id="734e0-153">В примере кода ниже показано, как получить ячейку, расположенную в строке 2 и столбце 5 листа **Sample** (Пример), загрузить его свойства `address` и `values` и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="734e0-153">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its `address` and `values` properties, and writes a message to the console.</span></span> <span data-ttu-id="734e0-154">Значения, передаваемые в метод `getCell(row: number, column:number)`, представляют собой индексируемые с нуля номера строк и столбцов получаемой ячейки.</span><span class="sxs-lookup"><span data-stu-id="734e0-154">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

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

## <a name="detect-data-changes"></a><span data-ttu-id="734e0-155">Обнаружение изменений данных</span><span class="sxs-lookup"><span data-stu-id="734e0-155">Detect data changes</span></span>

<span data-ttu-id="734e0-156">Возможно, надстройке потребуется реагировать на изменения пользователями данных в листе.</span><span class="sxs-lookup"><span data-stu-id="734e0-156">Your add-in may need to react to users changing the data in a worksheet.</span></span> <span data-ttu-id="734e0-157">Чтобы обнаружить эти изменения, можно [зарегистрировать обработчик событий](excel-add-ins-events.md#register-an-event-handler) для события `onChanged` листа.</span><span class="sxs-lookup"><span data-stu-id="734e0-157">To detect these changes, you can [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a worksheet.</span></span> <span data-ttu-id="734e0-158">Обработчики события `onChanged` получают объект [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) при возникновении события.</span><span class="sxs-lookup"><span data-stu-id="734e0-158">Event handlers for the `onChanged` event receive a [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) object when the event fires.</span></span>

<span data-ttu-id="734e0-159">Объект `WorksheetChangedEventArgs` предоставляет сведения об изменениях и источнике.</span><span class="sxs-lookup"><span data-stu-id="734e0-159">The `WorksheetChangedEventArgs` object provides information about the changes and the source.</span></span> <span data-ttu-id="734e0-160">Так как событие `onChanged` возникает при изменении формата или значения данных, может быть полезно, чтобы надстройка проверяла, действительно ли значения изменились.</span><span class="sxs-lookup"><span data-stu-id="734e0-160">Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values have actually changed.</span></span> <span data-ttu-id="734e0-161">Свойство `details` объединяет эти сведения в виде интерфейса [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span><span class="sxs-lookup"><span data-stu-id="734e0-161">The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span></span> <span data-ttu-id="734e0-162">В следующем примере кода показано, как отобразить значения и типы измененной ячейки до и после изменения.</span><span class="sxs-lookup"><span data-stu-id="734e0-162">The following code sample shows how to display the before and after values and types of a cell that has been changed.</span></span>

```js
// This function would be used as an event handler for the Worksheet.onChanged event.
function onWorksheetChanged(eventArgs) {
    Excel.run(function (context) {
        var details = eventArgs.details;
        var address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="handle-sorting-events"></a><span data-ttu-id="734e0-163">Обработка событий сортировки</span><span class="sxs-lookup"><span data-stu-id="734e0-163">Handle sorting events</span></span>

<span data-ttu-id="734e0-164">События `onColumnSorted` и `onRowSorted` указывают на сортировку любых данных на листе.</span><span class="sxs-lookup"><span data-stu-id="734e0-164">The `onColumnSorted` and `onRowSorted` events indicate when any worksheet data is sorted.</span></span> <span data-ttu-id="734e0-165">Эти события связаны с индивидуальными объектами `Worksheet` и с `WorkbookCollection` книги.</span><span class="sxs-lookup"><span data-stu-id="734e0-165">These events are connected to individual `Worksheet` objects and to the workbook's `WorkbookCollection`.</span></span> <span data-ttu-id="734e0-166">Они срабатывают при выполнении сортировки (программным образом или вручную с помощью пользовательского интерфейса Excel).</span><span class="sxs-lookup"><span data-stu-id="734e0-166">They fire whether the sorting is done programmatically or manually through the Excel user interface.</span></span>

> [!NOTE]
> <span data-ttu-id="734e0-167">Событие `onColumnSorted` срабатывает при сортировке столбцов в результате операции сортировки слева направо.</span><span class="sxs-lookup"><span data-stu-id="734e0-167">`onColumnSorted` fires when columns are sorted as the result of a left-to-right sort operation.</span></span> <span data-ttu-id="734e0-168">Событие `onRowSorted` срабатывает при сортировке строк в результате операции сортировки сверху вниз.</span><span class="sxs-lookup"><span data-stu-id="734e0-168">`onRowSorted` fires when rows are sorted as the result of a top-to-bottom sort operation.</span></span> <span data-ttu-id="734e0-169">Сортировка таблицы с помощью раскрывающегося меню в заголовке столбца приводит к срабатыванию события `onRowSorted`.</span><span class="sxs-lookup"><span data-stu-id="734e0-169">Sorting a table using the drop-down menu on a column header results in an `onRowSorted` event.</span></span> <span data-ttu-id="734e0-170">Событие соответствует перемещаемым данным, а не критериям сортировки.</span><span class="sxs-lookup"><span data-stu-id="734e0-170">The event corresponds with what is moving, not what is being considered as the sorting criteria.</span></span>

<span data-ttu-id="734e0-171">События `onColumnSorted` и `onRowSorted` реализуют функции обратного вызова соответственно с помощью объектов [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) и [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs).</span><span class="sxs-lookup"><span data-stu-id="734e0-171">The `onColumnSorted` and `onRowSorted` events provide their callbacks with [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) or [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs), respectively.</span></span> <span data-ttu-id="734e0-172">Эти объекты предоставляют более подробную информацию о событии.</span><span class="sxs-lookup"><span data-stu-id="734e0-172">These give more details about the event.</span></span> <span data-ttu-id="734e0-173">В частности, оба `EventArgs` обладают свойством `address`, которое представляет строки или столбцы, перемещенные в результате операции сортировки.</span><span class="sxs-lookup"><span data-stu-id="734e0-173">In particular, both `EventArgs` have an `address` property that represents the rows or columns moved as a result of the sort operation.</span></span> <span data-ttu-id="734e0-174">Включаются все ячейки, содержимое которых было отсортировано, даже если значение ячейки не входит в состав критериев сортировки.</span><span class="sxs-lookup"><span data-stu-id="734e0-174">Any cell with sorted content is included, even if that cell's value was not part of the sorting criteria.</span></span>

<span data-ttu-id="734e0-175">На приведенных ниже рисунках показаны диапазоны, возвращенные свойством `address` для событий сортировки.</span><span class="sxs-lookup"><span data-stu-id="734e0-175">The following images show the ranges returned by the `address` property for sort events.</span></span> <span data-ttu-id="734e0-176">Вот образец данных до сортировки:</span><span class="sxs-lookup"><span data-stu-id="734e0-176">First, here is the sample data before sorting:</span></span>

![Данные из таблицы в Excel до сортировки](../images/excel-sort-event-before.png)

<span data-ttu-id="734e0-178&quot;>Если выполнить сортировку сверху вниз для &quot;**Q1**&quot; (значения в &quot;**B**"), `WorksheetRowSortedEventArgs.address` возвращает следующие выделенные строки:</span><span class="sxs-lookup"><span data-stu-id="734e0-178">If a top-to-bottom sort is performed on "**Q1**" (the values in "**B**"), the following highlighted rows are returned by `WorksheetRowSortedEventArgs.address`:</span></span>

![Данные из таблицы в Excel после сортировки сверху вниз.](../images/excel-sort-event-after-row.png)

<span data-ttu-id="734e0-181&quot;>Если выполнить сортировку слева направо для &quot;**Quinces**&quot; (значения в &quot;**4**") в исходных данных, `WorksheetColumnsSortedEventArgs.address` возвращает следующие выделенные столбцы:</span><span class="sxs-lookup"><span data-stu-id="734e0-181">If a left-to-right sort is performed on "**Quinces**" (the values in "**4**") on the original data, the following highlighted columns are returned by `WorksheetColumnsSortedEventArgs.address`:</span></span>

![Данные из таблицы в Excel после сортировки слева направо.](../images/excel-sort-event-after-column.png)

<span data-ttu-id="734e0-184">В приведенном ниже примере кода показано, как зарегистрировать обработчик событий для события `Worksheet.onRowSorted`.</span><span class="sxs-lookup"><span data-stu-id="734e0-184">The following code sample shows how to register an event handler for the `Worksheet.onRowSorted` event.</span></span> <span data-ttu-id="734e0-185">Обратный вызов обработчика очищает цвет заливки для диапазона, затем применяет заливку к ячейкам перемещенных строк.</span><span class="sxs-lookup"><span data-stu-id="734e0-185">The handler's callback clears the fill color for the range, then fills the cells of the moved rows.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // This will fire whenever a row has been moved as the result of a sort action.
    sheet.onRowSorted.add(function (event) {
        return Excel.run(function (context) {
            console.log("Row sorted: " + event.address);
            var sheet = context.workbook.worksheets.getActiveWorksheet();

            // Clear formatting for section, then highlight the sorted area.
            sheet.getRange("A1:E5").format.fill.clear();
            if (event.address !== "") {
                sheet.getRanges(event.address).format.fill.color = "yellow";
            }

            return context.sync();
        });
    });

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="find-all-cells-with-matching-text"></a><span data-ttu-id="734e0-186">Поиск всех ячеек с соответствующим текстом</span><span class="sxs-lookup"><span data-stu-id="734e0-186">Find all cells with matching text</span></span>

<span data-ttu-id="734e0-187">У объекта `Worksheet` есть метод `find` для поиска указанной строки в листе.</span><span class="sxs-lookup"><span data-stu-id="734e0-187">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="734e0-188">Он возвращает объект `RangeAreas`, являющийся коллекцией объектов `Range`, которые можно отредактировать все сразу.</span><span class="sxs-lookup"><span data-stu-id="734e0-188">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="734e0-189">Приведенный ниже пример кода находит все ячейки со значениями, соответствующими строке **Complete** (Завершено), и окрашивает их зеленым цветом.</span><span class="sxs-lookup"><span data-stu-id="734e0-189">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="734e0-190">Обратите внимание, что метод `findAll` выдаст ошибку `ItemNotFound`, если указанной строки не существует в листе.</span><span class="sxs-lookup"><span data-stu-id="734e0-190">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="734e0-191">Если ожидается, что указанная строка может отсутствовать в листе, используйте вместо этого метод [findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), чтобы ваш код корректно обработал этот сценарий.</span><span class="sxs-lookup"><span data-stu-id="734e0-191">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // findAll will match the whole cell value
        matchCase: false // findAll will not match case
    });

    return context.sync()
        .then(function() {
            foundRanges.format.fill.color = "green"
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="734e0-192">В этом разделе описано, как найти ячейки и диапазоны с помощью функций объекта `Worksheet`.</span><span class="sxs-lookup"><span data-stu-id="734e0-192">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="734e0-193">Дополнительные сведения об извлечении диапазонов можно найти в статьях о конкретных объектах.</span><span class="sxs-lookup"><span data-stu-id="734e0-193">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="734e0-194">Примеры получения диапазона в листах с помощью объекта см. в примере `Range` Get a range using the Excel [JavaScript API.](excel-add-ins-ranges-get.md)</span><span class="sxs-lookup"><span data-stu-id="734e0-194">For examples that show how to get a range within a worksheet using the `Range` object, see [Get a range using the Excel JavaScript API](excel-add-ins-ranges-get.md).</span></span>
> - <span data-ttu-id="734e0-195">Примеры, в которых показано, как получить диапазоны из объекта `Table`, см. в статье [Работа с таблицами с использованием API JavaScript для Excel](excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="734e0-195">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="734e0-196">Примеры, в которых показано, как выполнять поиск большого диапазона для нескольких поддиапазонов с учетом характеристик ячеек, см. в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="734e0-196">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="filter-data"></a><span data-ttu-id="734e0-197">Фильтрация данных</span><span class="sxs-lookup"><span data-stu-id="734e0-197">Filter data</span></span>

<span data-ttu-id="734e0-198">Объект [AutoFilter](/javascript/api/excel/excel.autofilter) применяет фильтры данных в диапазоне на листе.</span><span class="sxs-lookup"><span data-stu-id="734e0-198">An [AutoFilter](/javascript/api/excel/excel.autofilter) applies data filters across a range within the worksheet.</span></span> <span data-ttu-id="734e0-199">Он создается с помощью метода `Worksheet.autoFilter.apply`, содержащего следующие параметры:</span><span class="sxs-lookup"><span data-stu-id="734e0-199">This is created with `Worksheet.autoFilter.apply`, which has the following parameters:</span></span>

- <span data-ttu-id="734e0-200">`range`: диапазон, к которому применяется фильтр, указанный в виде объекта `Range` или строки.</span><span class="sxs-lookup"><span data-stu-id="734e0-200">`range`: The range to which the filter is applied, specified as either a `Range` object or a string.</span></span>
- <span data-ttu-id="734e0-201">`columnIndex`: отсчитываемый от нуля индекс столбца, по которому оценивается условие фильтра.</span><span class="sxs-lookup"><span data-stu-id="734e0-201">`columnIndex`: The zero-based column index against which the filter criteria is evaluated.</span></span>
- <span data-ttu-id="734e0-202">`criteria`: объект [FilterCriteria](/javascript/api/excel/excel.filtercriteria), определяющий, какие строки следует фильтровать на основе ячейки столбца.</span><span class="sxs-lookup"><span data-stu-id="734e0-202">`criteria`: A [FilterCriteria](/javascript/api/excel/excel.filtercriteria) object determining which rows should be filtered based on the column's cell.</span></span>

<span data-ttu-id="734e0-203">В первом примере кода показано, как добавить фильтр в используемый диапазон на листе.</span><span class="sxs-lookup"><span data-stu-id="734e0-203">The first code sample shows how to add a filter to the worksheet's used range.</span></span> <span data-ttu-id="734e0-204">Этот фильтр скрывает записи, не входящие в верхние 25 %, на основе значений в столбце **3**.</span><span class="sxs-lookup"><span data-stu-id="734e0-204">This filter will hide entries that are not in the top 25%, based on the values in column **3**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var farmData = sheet.getUsedRange();

    // This filter will only show the rows with the top 25% of values in column 3.
    sheet.autoFilter.apply(farmData, 3, { criterion1: "25", filterOn: Excel.FilterOn.topPercent });
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="734e0-205">В следующем примере кода показано, как обновить автофильтр, используя метод `reapply`.</span><span class="sxs-lookup"><span data-stu-id="734e0-205">The next code sample shows how to refresh the auto-filter using the `reapply` method.</span></span> <span data-ttu-id="734e0-206">Это следует выполнять при изменении данных в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="734e0-206">This should be done when the data in the range changes.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.reapply();
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="734e0-207">В последнем примере кода автофильтра показано, как удалить автофильтр с листа с помощью метода `remove`.</span><span class="sxs-lookup"><span data-stu-id="734e0-207">The final auto-filter code sample shows how to remove the auto-filter from the worksheet with the `remove` method.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="734e0-208">Объект `AutoFilter` также можно применять к отдельным таблицам.</span><span class="sxs-lookup"><span data-stu-id="734e0-208">An `AutoFilter` can also be applied to individual tables.</span></span> <span data-ttu-id="734e0-209">Дополнительные сведения см. в статье [Работа с таблицами с использованием API JavaScript для Excel](excel-add-ins-tables.md#autofilter).</span><span class="sxs-lookup"><span data-stu-id="734e0-209">See [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md#autofilter) for more information.</span></span>

## <a name="data-protection"></a><span data-ttu-id="734e0-210">Защита данных</span><span class="sxs-lookup"><span data-stu-id="734e0-210">Data protection</span></span>

<span data-ttu-id="734e0-211">Надстройка может управлять возможностью пользователя по изменению данных на листе.</span><span class="sxs-lookup"><span data-stu-id="734e0-211">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="734e0-212">Свойство `protection` листа является объектом [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) с методом `protect()`.</span><span class="sxs-lookup"><span data-stu-id="734e0-212">The worksheet's `protection` property is a [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="734e0-213">В приведенном ниже примере показан основной сценарий переключения полной защиты активного листа.</span><span class="sxs-lookup"><span data-stu-id="734e0-213">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("protection/protected");

    return context.sync().then(function() {
        if (!activeSheet.protection.protected) {
            activeSheet.protection.protect();
        }
    })
}).catch(errorHandlerFunction);
```

<span data-ttu-id="734e0-214">Метод `protect` содержит два необязательных параметра:</span><span class="sxs-lookup"><span data-stu-id="734e0-214">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="734e0-215">`options`: объект [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions), определяющий конкретные ограничения на редактирование.</span><span class="sxs-lookup"><span data-stu-id="734e0-215">`options`: A [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="734e0-216">`password`: строка, представляющая пароль, необходимый пользователю для обхода защиты и редактирования листа.</span><span class="sxs-lookup"><span data-stu-id="734e0-216">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="734e0-217">В статье [Защита листа](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) содержатся дополнительные сведения о защите листа и ее изменении с помощью пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="734e0-217">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="page-layout-and-print-settings"></a><span data-ttu-id="734e0-218">Параметры разметки страницы и печати</span><span class="sxs-lookup"><span data-stu-id="734e0-218">Page layout and print settings</span></span>

<span data-ttu-id="734e0-219">Надстройкам доступны параметры разметки страницы на уровне листа.</span><span class="sxs-lookup"><span data-stu-id="734e0-219">Add-ins have access to page layout settings at a worksheet level.</span></span> <span data-ttu-id="734e0-220">Они управляют печатью листа.</span><span class="sxs-lookup"><span data-stu-id="734e0-220">These control how the sheet is printed.</span></span> <span data-ttu-id="734e0-221">У объекта `Worksheet` есть три связанных с разметкой свойства: `horizontalPageBreaks`, `verticalPageBreaks`, `pageLayout`.</span><span class="sxs-lookup"><span data-stu-id="734e0-221">A `Worksheet` object has three layout-related properties: `horizontalPageBreaks`, `verticalPageBreaks`, `pageLayout`.</span></span>

<span data-ttu-id="734e0-222">`Worksheet.horizontalPageBreaks` и `Worksheet.verticalPageBreaks` относятся к [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection).</span><span class="sxs-lookup"><span data-stu-id="734e0-222">`Worksheet.horizontalPageBreaks` and `Worksheet.verticalPageBreaks` are [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection).</span></span> <span data-ttu-id="734e0-223">Это коллекции объектов [PageBreak](/javascript/api/excel/excel.pagebreak), указывающих диапазоны вставки разрывов страниц, добавляемых вручную.</span><span class="sxs-lookup"><span data-stu-id="734e0-223">These are collections of [PageBreaks](/javascript/api/excel/excel.pagebreak), which specify ranges where manual page breaks are inserted.</span></span> <span data-ttu-id="734e0-224">В следующем примере кода добавляется горизонтальный разрыв страницы над строкой **21**.</span><span class="sxs-lookup"><span data-stu-id="734e0-224">The following code sample adds a horizontal page break above row **21**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break is added above this range.
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="734e0-225">`Worksheet.pageLayout` является объектом [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="734e0-225">`Worksheet.pageLayout` is a [PageLayout](/javascript/api/excel/excel.pagelayout) object.</span></span> <span data-ttu-id="734e0-226">Этот объект содержит параметры разметки и печати, не зависящие от применения конкретного принтера.</span><span class="sxs-lookup"><span data-stu-id="734e0-226">This object contains layout and print settings that are not dependent any printer-specific implementation.</span></span> <span data-ttu-id="734e0-227">Эти параметры включают поля, ориентацию, нумерацию страницы, строки заголовков и область печати.</span><span class="sxs-lookup"><span data-stu-id="734e0-227">These settings include margins, orientation, page numbering, title rows, and print area.</span></span>

<span data-ttu-id="734e0-228">В следующем примере кода страница выравнивается по центру (по вертикали и горизонтали), устанавливается строка заголовка, которая печатается в верхней части каждой страницы, и задается подраздел листа в качестве области печати.</span><span class="sxs-lookup"><span data-stu-id="734e0-228">The following code sample centers the page (both vertically and horizontally), sets a title row that will be printed at the top of every page, and sets the printed area to a subsection of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the area to be printed to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="734e0-229">См. также</span><span class="sxs-lookup"><span data-stu-id="734e0-229">See also</span></span>

- [<span data-ttu-id="734e0-230">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="734e0-230">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
