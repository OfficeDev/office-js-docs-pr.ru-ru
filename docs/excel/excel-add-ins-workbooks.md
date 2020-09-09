---
title: Работа с книгами с использованием API JavaScript для Excel
description: Примеры кода, в которых показано, как выполнять распространенные задачи с книгами или функциями уровня приложения с помощью API JavaScript для Excel.
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 1da664b5ebbc8bd44b1c4e558a48997291e8967e
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408700"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="18828-103">Работа с книгами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="18828-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="18828-104">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для книг с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="18828-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="18828-105">Полный список свойств и методов, `Workbook` поддерживаемых объектом, представлен в статье [объект WORKBOOK (API JavaScript для Excel)](/javascript/api/excel/excel.workbook).</span><span class="sxs-lookup"><span data-stu-id="18828-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="18828-106">В этой статье также рассматриваются действия на уровне книги, выполняемые с помощью объекта [Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="18828-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="18828-107">Объект Workbook — это точка входа для вашей надстройки для взаимодействия с Excel.</span><span class="sxs-lookup"><span data-stu-id="18828-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="18828-108">Он поддерживает коллекции листов, таблиц, сводных таблиц и других элементов, через которые выполняется доступ и изменение данных Excel.</span><span class="sxs-lookup"><span data-stu-id="18828-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="18828-109">Объект [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) предоставляет надстройке доступ ко всем данным книги с помощью отдельных листов.</span><span class="sxs-lookup"><span data-stu-id="18828-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="18828-110">В частности, он позволяет надстройке добавлять листы, перемещаться между ними и назначать обработчиков событий листа.</span><span class="sxs-lookup"><span data-stu-id="18828-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="18828-111">В статье [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md) описывается способ доступа к листам и их изменение.</span><span class="sxs-lookup"><span data-stu-id="18828-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="18828-112">Получение активной ячейки или выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="18828-112">Get the active cell or selected range</span></span>

<span data-ttu-id="18828-113">Объект Workbook содержит два метода для получения диапазона ячеек, выделенных пользователем или надстройкой: `getActiveCell()` и `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="18828-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="18828-114">`getActiveCell()` получает активную ячейку из книги в виде [объекта Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="18828-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="18828-115">В приведенном ниже примере показан вызов `getActiveCell()` с последующей печатью адреса ячейки в консоль.</span><span class="sxs-lookup"><span data-stu-id="18828-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="18828-116">Метод `getSelectedRange()` возвращает один диапазон, выделенный в настоящее время.</span><span class="sxs-lookup"><span data-stu-id="18828-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="18828-117">Если выделено несколько диапазонов, возникает ошибка InvalidSelection.</span><span class="sxs-lookup"><span data-stu-id="18828-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="18828-118">В приведенном ниже примере показан вызов метода `getSelectedRange()`, который затем устанавливает желтый цвет заливки для диапазона.</span><span class="sxs-lookup"><span data-stu-id="18828-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="18828-119">Создание книги</span><span class="sxs-lookup"><span data-stu-id="18828-119">Create a workbook</span></span>

<span data-ttu-id="18828-120">Ваша надстройка может создать новую книгу, отдельную от экземпляра Excel, в котором в настоящее время работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="18828-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="18828-121">Для этой цели в объекте Excel имеется метод `createWorkbook`.</span><span class="sxs-lookup"><span data-stu-id="18828-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="18828-122">При вызове этого метода сразу открывается и отображается новая книга в новом экземпляре программы Excel.</span><span class="sxs-lookup"><span data-stu-id="18828-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="18828-123">Ваша надстройка остается открытой и запущенной в предыдущей книге.</span><span class="sxs-lookup"><span data-stu-id="18828-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="18828-124">С помощью метода `createWorkbook` также можно создать копию существующей книги.</span><span class="sxs-lookup"><span data-stu-id="18828-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="18828-125">Метод принимает в качестве необязательного параметра строковое представление XLSX-файла в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="18828-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="18828-126">Полученная книга будет копией этого файла, предполагая, что строковый аргумент является допустимым XLSX-файлом.</span><span class="sxs-lookup"><span data-stu-id="18828-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="18828-127">Вы можете получить текущую книгу надстройки в виде строки в кодировке Base64 с помощью [фрагментирования файлов](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="18828-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="18828-128">Преобразование файла в нужную строку в кодировке base64 можно выполнить с помощью класса [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader), как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="18828-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="18828-129">Вставьте копию существующей книги в текущую (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="18828-129">Insert a copy of an existing workbook into the current one (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="18828-130">В настоящее время метод `WorksheetCollection.addFromBase64` доступен только в общедоступной предварительной версии и только в Office для Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="18828-130">The `WorksheetCollection.addFromBase64` method is currently only available in public preview and only for Office on Windows and Mac.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="18828-131">В предыдущем примере показана новая книга, которая была создана из существующей книги.</span><span class="sxs-lookup"><span data-stu-id="18828-131">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="18828-132">Вы также можете скопировать отдельные части или всю существующую книгу целиком в книгу, привязанную в настоящее время к вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="18828-132">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="18828-133">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) для книги имеет метод `addFromBase64` для вставки копий листов целевой книги в саму книгу.</span><span class="sxs-lookup"><span data-stu-id="18828-133">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="18828-134">Файл другой книги передается в виде строки в кодировке base64, как и вызов `Excel.createWorkbook`.</span><span class="sxs-lookup"><span data-stu-id="18828-134">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="18828-135">В примере ниже показаны листы книги, которые были вставлены в текущую книгу непосредственно после активного листа.</span><span class="sxs-lookup"><span data-stu-id="18828-135">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="18828-136">Обратите внимание, что `null` передается для параметра `sheetNamesToInsert?: string[]`.</span><span class="sxs-lookup"><span data-stu-id="18828-136">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="18828-137">Это означает, что все листы были вставлены.</span><span class="sxs-lookup"><span data-stu-id="18828-137">This means all the worksheets are being inserted.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        var sheets = context.workbook.worksheets;
        sheets.addFromBase64(
            workbookContents,
            null, // get all the worksheets
            Excel.WorksheetPositionType.after, // insert them after the worksheet specified by the next parameter
            sheets.getActiveWorksheet() // insert them after the active worksheet
        );
        return context.sync();
    });
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="18828-138">Защита структуры книги</span><span class="sxs-lookup"><span data-stu-id="18828-138">Protect the workbook's structure</span></span>

<span data-ttu-id="18828-139">Надстройка может управлять возможностью пользователя по изменению структуры книги.</span><span class="sxs-lookup"><span data-stu-id="18828-139">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="18828-140">Свойство `protection` объекта Workbook является объектом [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) с методом `protect()`.</span><span class="sxs-lookup"><span data-stu-id="18828-140">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="18828-141">В приведенном ниже примере показан основной сценарий переключения защиты структуры книги.</span><span class="sxs-lookup"><span data-stu-id="18828-141">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

```js
Excel.run(function (context) {
    var workbook = context.workbook;
    workbook.load("protection/protected");

    return context.sync().then(function() {
        if (!workbook.protection.protected) {
            workbook.protection.protect();
        }
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="18828-142">Метод `protect` принимает необязательный строковый параметр.</span><span class="sxs-lookup"><span data-stu-id="18828-142">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="18828-143">Эта строка представляет пароль, необходимый пользователю для обхода защиты и изменения структуры книги.</span><span class="sxs-lookup"><span data-stu-id="18828-143">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="18828-144">Защиту также можно установить на уровне книги, чтобы предотвратить нежелательные изменения данных.</span><span class="sxs-lookup"><span data-stu-id="18828-144">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="18828-145">Дополнительные сведения см. в разделе **Защита данных** статьи [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="18828-145">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="18828-146">Дополнительные сведения о защите книги в Excel см. в статье [Защита книги](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="18828-146">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="18828-147">Доступ к свойствам документов</span><span class="sxs-lookup"><span data-stu-id="18828-147">Access document properties</span></span>

<span data-ttu-id="18828-148">Объекты Workbook имеют доступ к метаданным файлов Office, называемым [свойствами документов](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="18828-148">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="18828-149">Свойство `properties` объекта Workbook является объектом [DocumentProperties](/javascript/api/excel/excel.documentproperties), содержащим эти значения метаданных.</span><span class="sxs-lookup"><span data-stu-id="18828-149">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="18828-150">В приведенном ниже примере показано, как задать `author` свойство.</span><span class="sxs-lookup"><span data-stu-id="18828-150">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a><span data-ttu-id="18828-151">Настраиваемые свойства</span><span class="sxs-lookup"><span data-stu-id="18828-151">Custom properties</span></span>

<span data-ttu-id="18828-152">Также можно установить настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="18828-152">You can also define custom properties.</span></span> <span data-ttu-id="18828-153">Объект DocumentProperties содержит свойство `custom`, представляющее коллекцию пар "ключ-значение" для свойств, определяемых пользователем.</span><span class="sxs-lookup"><span data-stu-id="18828-153">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="18828-154">В приведенном ниже примере показано, как создать настраиваемое свойство с именем **Introduction** со значением "Hello", а затем вызвать его.</span><span class="sxs-lookup"><span data-stu-id="18828-154">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

```js
Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    var customProperty = customDocProperties.getItem("Introduction");
    customProperty.load(["key, value"]);

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

#### <a name="worksheet-level-custom-properties-preview"></a><span data-ttu-id="18828-155">Настраиваемые свойства на уровне листа (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="18828-155">Worksheet-level custom properties (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="18828-156">Настраиваемые свойства на уровне листа в настоящее время находятся в режиме предварительного просмотра.</span><span class="sxs-lookup"><span data-stu-id="18828-156">Worksheet-level custom properties are currently in preview.</span></span> [!INCLUDE [Information about using preview Excel APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="18828-157">Настраиваемые свойства также можно задать на уровне листа.</span><span class="sxs-lookup"><span data-stu-id="18828-157">Custom properties can also be set at the worksheet level.</span></span> <span data-ttu-id="18828-158">Они похожи на настраиваемые свойства на уровне документа, за исключением того, что один и тот же ключ может повторяться на разных листах.</span><span class="sxs-lookup"><span data-stu-id="18828-158">These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.</span></span> <span data-ttu-id="18828-159">В приведенном ниже примере показано, как создать настраиваемое свойство с именем **воркшитграуп** со значением "Alpha" на текущем листе, а затем извлечь его.</span><span class="sxs-lookup"><span data-stu-id="18828-159">The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.</span></span>

```js
Excel.run(function (context) {
    // Add the custom property.
    var customWorksheetProperties = context.workbook.worksheets.getActiveWorksheet().customProperties;
    customWorksheetProperties.add("WorksheetGroup", "Alpha");

    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    // Load the keys and values of all custom properties in the current worksheet.
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    var customWorksheetProperties = worksheet.customProperties;
    var customWorksheetProperty = customWorksheetProperties.getItem("WorksheetGroup");
    customWorksheetProperty.load(["key", "value"]);

    return context.sync().then(function() {
        // Log the WorksheetGroup custom property to the console.
        console.log(worksheet.name + ": " + customWorksheetProperty.key); // "WorksheetGroup"
        console.log("  Custom value : " + customWorksheetProperty.value); // "Alpha"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="18828-160">Доступ к параметрам документа</span><span class="sxs-lookup"><span data-stu-id="18828-160">Access document settings</span></span>

<span data-ttu-id="18828-161">Параметры книги похожи на коллекцию настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="18828-161">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="18828-162">Различие заключается в том, что параметры уникальны для одного файла Excel и соединения надстройки, а свойства связаны только с файлом.</span><span class="sxs-lookup"><span data-stu-id="18828-162">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="18828-163">В приведенном ниже примере показано, как создать параметр и получить к нему доступ.</span><span class="sxs-lookup"><span data-stu-id="18828-163">The following example shows how to create and access a setting.</span></span>

```js
Excel.run(function (context) {
    var settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    var needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    return context.sync().then(function() {
        console.log("Workbook needs review : " + needsReview.value);
    });
}).catch(errorHandlerFunction);
```

## <a name="access-application-culture-settings"></a><span data-ttu-id="18828-164">Параметры культуры приложения Access</span><span class="sxs-lookup"><span data-stu-id="18828-164">Access application culture settings</span></span>

<span data-ttu-id="18828-165">Книга содержит параметры языка и региональных параметров, которые влияют на отображение определенных данных.</span><span class="sxs-lookup"><span data-stu-id="18828-165">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="18828-166">Эти параметры могут помочь локализовать данные, когда пользователи надстройки совместно работают с книгами на различных языках и региональных параметрах.</span><span class="sxs-lookup"><span data-stu-id="18828-166">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="18828-167">Надстройка может использовать синтаксический анализ строк для локализации формата чисел, дат и времени на основе параметров языковых параметров системы, чтобы каждый пользователь видел данные в формате языка и региональных параметров.</span><span class="sxs-lookup"><span data-stu-id="18828-167">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="18828-168">`Application.cultureInfo` Определяет параметры языка и региональных параметров системы в виде объекта [CultureInfo](/javascript/api/excel/excel.cultureinfo) .</span><span class="sxs-lookup"><span data-stu-id="18828-168">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="18828-169">Содержит такие параметры, как числовой десятичный разделитель или формат даты.</span><span class="sxs-lookup"><span data-stu-id="18828-169">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="18828-170">Некоторые параметры культуры можно [изменить с помощью пользовательского интерфейса Excel](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span><span class="sxs-lookup"><span data-stu-id="18828-170">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="18828-171">Параметры системы сохраняются в `CultureInfo` объекте.</span><span class="sxs-lookup"><span data-stu-id="18828-171">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="18828-172">Все локальные изменения хранятся в виде свойств уровня [приложения](/javascript/api/excel/excel.application), например `Application.decimalSeparator` .</span><span class="sxs-lookup"><span data-stu-id="18828-172">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="18828-173">В примере ниже показано, как изменить символ десятичного разделителя в числовой строке с "," на символ, используемый параметрами системы.</span><span class="sxs-lookup"><span data-stu-id="18828-173">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var decimalSource = sheet.getRange("B2");
    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");

    return context.sync().then(function() {
        var systemDecimalSeparator =
            context.application.cultureInfo.numberFormat.numberDecimalSeparator;
        var oldDecimalString = decimalSource.values[0][0];

        // This assumes the input column is standardized to use "," as the decimal separator.
        var newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

        var resultRange = sheet.getRange("C2");
        resultRange.values = [[newDecimalString]];
        resultRange.format.autofitColumns();
        return context.sync();
    });
});
```

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="18828-174">Добавление настраиваемых XML-данных в книгу</span><span class="sxs-lookup"><span data-stu-id="18828-174">Add custom XML data to the workbook</span></span>

<span data-ttu-id="18828-175">Формат файла Excel Open XML **(XLSX)** позволяет надстройке внедрить настраиваемые XML-данные в книгу.</span><span class="sxs-lookup"><span data-stu-id="18828-175">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="18828-176">Эти данные сохраняются с книгой независимо от надстройки.</span><span class="sxs-lookup"><span data-stu-id="18828-176">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="18828-177">Книга содержит объект [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), являющийся списком объектов [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="18828-177">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="18828-178">Они предоставляют доступ к строкам XML и соответствующему уникальному идентификатору.</span><span class="sxs-lookup"><span data-stu-id="18828-178">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="18828-179">Сохраняя эти идентификаторы как параметры, надстройка может сохранять ключи к частям XML между сеансами.</span><span class="sxs-lookup"><span data-stu-id="18828-179">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="18828-180">В приведенных ниже примерах показано, как использовать настраиваемые части XML.</span><span class="sxs-lookup"><span data-stu-id="18828-180">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="18828-181">В первом блоке кода показано, как внедрять XML-данные в документ.</span><span class="sxs-lookup"><span data-stu-id="18828-181">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="18828-182">Выполняется сохранение списка проверяющих, а затем используются параметры книги, чтобы сохранить параметр `id` XML для будущих извлечений.</span><span class="sxs-lookup"><span data-stu-id="18828-182">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="18828-183">Во втором блоке показано, как получить доступ к этим XML-данным позднее.</span><span class="sxs-lookup"><span data-stu-id="18828-183">The second block shows how to access that XML later.</span></span> <span data-ttu-id="18828-184">Параметр "ContosoReviewXmlPartId" загружается и передается объекту `customXmlParts` книги.</span><span class="sxs-lookup"><span data-stu-id="18828-184">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="18828-185">Данные XML затем печатаются в консоль.</span><span class="sxs-lookup"><span data-stu-id="18828-185">The XML data is then printed to the console.</span></span>

```js
Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    var originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    var customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");

    return context.sync().then(function() {
        // Store the XML part's ID in a setting
        var settings = context.workbook.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);
    });
}).catch(errorHandlerFunction);
```

```js
Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    var settings = context.workbook.settings;
    var xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    return context.sync().then(function () {
        if (xmlPartIDSetting.value) {
            var customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
            var xmlBlob = customXmlPart.getXml();

            return context.sync().then(function () {
                // Add spaces to make more human readable in the console
                var readableXML = xmlBlob.value.replace(/></g, "> <");
                console.log(readableXML);
            });
        }
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="18828-186">`CustomXMLPart.namespaceUri` заполняется только в том случае, если настраиваемый XML-элемент верхнего уровня содержит атрибут `xmlns`.</span><span class="sxs-lookup"><span data-stu-id="18828-186">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="18828-187">Управление режимом вычислений</span><span class="sxs-lookup"><span data-stu-id="18828-187">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="18828-188">Установка режима вычислений</span><span class="sxs-lookup"><span data-stu-id="18828-188">Set calculation mode</span></span>

<span data-ttu-id="18828-189">По умолчанию Excel пересчитывает результаты формул при каждом изменении ячейки из ссылки.</span><span class="sxs-lookup"><span data-stu-id="18828-189">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="18828-190">Производительность вашей надстройки можно улучшить путем изменения режима вычислений.</span><span class="sxs-lookup"><span data-stu-id="18828-190">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="18828-191">У объекта Application есть свойство `calculationMode` типа `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="18828-191">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="18828-192">Ему можно присвоить следующие значения:</span><span class="sxs-lookup"><span data-stu-id="18828-192">It can be set to the following values:</span></span>

- <span data-ttu-id="18828-193">`automatic`: режим пересчета по умолчанию, при котором Excel вычисляет новые результаты формулы при каждом изменении соответствующих данных.</span><span class="sxs-lookup"><span data-stu-id="18828-193">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="18828-194">`automaticExceptTables`: аналогично `automatic`, за исключением того, что игнорируются любые изменения значений таблиц.</span><span class="sxs-lookup"><span data-stu-id="18828-194">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="18828-195">`manual`: вычисления выполняются только в том случае, если пользователь или надстройка запрашивает их.</span><span class="sxs-lookup"><span data-stu-id="18828-195">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="18828-196">Установка типа вычислений</span><span class="sxs-lookup"><span data-stu-id="18828-196">Set calculation type</span></span>

<span data-ttu-id="18828-197">Объект [Application](/javascript/api/excel/excel.application) предоставляет метод применения немедленного пересчета.</span><span class="sxs-lookup"><span data-stu-id="18828-197">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="18828-198">Метод `Application.calculate(calculationType)` запускает ручной пересчет с учетом указанного типа `calculationType`.</span><span class="sxs-lookup"><span data-stu-id="18828-198">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="18828-199">Можно указать следующие значения:</span><span class="sxs-lookup"><span data-stu-id="18828-199">The following values can be specified:</span></span>

- <span data-ttu-id="18828-200">`full`: пересчет всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.</span><span class="sxs-lookup"><span data-stu-id="18828-200">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="18828-201">`fullRebuild`: проверка зависимых формул с последующим пересчетом всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.</span><span class="sxs-lookup"><span data-stu-id="18828-201">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="18828-202">`recalculate`: пересчет формул, которые были изменены (или помечены программным путем для пересчета) с момента последнего вычисления, и зависимых от них формул во всех активных книгах.</span><span class="sxs-lookup"><span data-stu-id="18828-202">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="18828-203">Дополнительные сведения о пересчете см. в статье [Изменение пересчета, итерации или точности формулы](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="18828-203">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="18828-204">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="18828-204">Temporarily suspend calculations</span></span>

<span data-ttu-id="18828-205">API Excel также позволяет надстройкам отключить вычисления до вызова `RequestContext.sync()`.</span><span class="sxs-lookup"><span data-stu-id="18828-205">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="18828-206">Для этого используется `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="18828-206">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="18828-207">Используйте этот метод, если ваша надстройка изменяет большие диапазоны без необходимости доступа к данным между изменениями.</span><span class="sxs-lookup"><span data-stu-id="18828-207">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="save-the-workbook"></a><span data-ttu-id="18828-208">Сохраните книгу.</span><span class="sxs-lookup"><span data-stu-id="18828-208">Save the workbook</span></span>

<span data-ttu-id="18828-209">`Workbook.save` сохраняет книгу в постоянное хранилище.</span><span class="sxs-lookup"><span data-stu-id="18828-209">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="18828-210">Метод `save` имеет один необязательный параметр `saveBehavior`, который может принимать одно из следующих значений:</span><span class="sxs-lookup"><span data-stu-id="18828-210">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="18828-211">`Excel.SaveBehavior.save` (по умолчанию): файл будет сохранен без предварительного запроса имени файла, а также место для сохранения.</span><span class="sxs-lookup"><span data-stu-id="18828-211">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="18828-212">Если файл не был сохранен ранее, он будет сохранен в папке по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="18828-212">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="18828-213">Если файл уже был сохранен ранее, он будет сохранен в той же папке.</span><span class="sxs-lookup"><span data-stu-id="18828-213">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="18828-214">`Excel.SaveBehavior.prompt`: если файл не был сохранен ранее, будет предложено ввести имя файла и место для сохранения.</span><span class="sxs-lookup"><span data-stu-id="18828-214">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="18828-215">Если файл уже был сохранен ранее, он будет сохраняться в той же папке, и никаких дополнительных действий не потребуется.</span><span class="sxs-lookup"><span data-stu-id="18828-215">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="18828-216">Если пользователь при запрос на сохранение отменяет операцию, `save` выдает исключение.</span><span class="sxs-lookup"><span data-stu-id="18828-216">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="18828-217">Закрытие книги.</span><span class="sxs-lookup"><span data-stu-id="18828-217">Close the workbook</span></span>

<span data-ttu-id="18828-218">`Workbook.close` закрывает книгу, а также надстройки, связанные с книгой, (приложение Excel остается открытым).</span><span class="sxs-lookup"><span data-stu-id="18828-218">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="18828-219">Метод `close` имеет один необязательный параметр `closeBehavior`, который может принимать одно из следующих значений:</span><span class="sxs-lookup"><span data-stu-id="18828-219">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="18828-220">`Excel.CloseBehavior.save` (по умолчанию): файл будет сохранен до закрытия.</span><span class="sxs-lookup"><span data-stu-id="18828-220">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="18828-221">Если файл не был сохранен ранее, будет предложено ввести имя файла и место для сохранения.</span><span class="sxs-lookup"><span data-stu-id="18828-221">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="18828-222">`Excel.CloseBehavior.skipSave`: файл будет немедленно закрыт без сохранения.</span><span class="sxs-lookup"><span data-stu-id="18828-222">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="18828-223">Все несохраненные изменения будут потеряны.</span><span class="sxs-lookup"><span data-stu-id="18828-223">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="18828-224">См. также</span><span class="sxs-lookup"><span data-stu-id="18828-224">See also</span></span>

- [<span data-ttu-id="18828-225">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="18828-225">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="18828-226">Работа с листами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="18828-226">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="18828-227">Работа с диапазонами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="18828-227">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
