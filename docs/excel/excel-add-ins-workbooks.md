---
title: Работа с книгами с использованием API JavaScript для Excel
description: Узнайте, как выполнять общие задачи с помощью книг или функций уровня приложений с помощью Excel API JavaScript.
ms.date: 06/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 78cdf57ce6ecce3e9e3e40188b3325cdf15ab265
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349429"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="2d7ce-103">Работа с книгами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2d7ce-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="2d7ce-104">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для книг с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="2d7ce-105">Полный список свойств и методов, поддерживаемых объектом, см. в книге `Workbook` [Объект (API JavaScript для Excel).](/javascript/api/excel/excel.workbook)</span><span class="sxs-lookup"><span data-stu-id="2d7ce-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="2d7ce-106">В этой статье также рассматриваются действия на уровне книги, выполняемые с помощью объекта [Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="2d7ce-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="2d7ce-107">Объект Workbook — это точка входа для вашей надстройки для взаимодействия с Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="2d7ce-108">Он поддерживает коллекции листов, таблиц, сводных таблиц и других элементов, через которые выполняется доступ и изменение данных Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="2d7ce-109">Объект [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) предоставляет надстройке доступ ко всем данным книги с помощью отдельных листов.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="2d7ce-110">В частности, он позволяет надстройке добавлять листы, перемещаться между ними и назначать обработчиков событий листа.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="2d7ce-111">В статье [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md) описывается способ доступа к листам и их изменение.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="2d7ce-112">Получение активной ячейки или выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="2d7ce-112">Get the active cell or selected range</span></span>

<span data-ttu-id="2d7ce-113">Объект Workbook содержит два метода для получения диапазона ячеек, выделенных пользователем или надстройкой: `getActiveCell()` и `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="2d7ce-114">`getActiveCell()` получает активную ячейку из книги в виде [объекта Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="2d7ce-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="2d7ce-115">В приведенном ниже примере показан вызов `getActiveCell()` с последующей печатью адреса ячейки в консоль.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="2d7ce-116">Метод `getSelectedRange()` возвращает один диапазон, выделенный в настоящее время.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="2d7ce-117">Если выделено несколько диапазонов, возникает ошибка InvalidSelection.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="2d7ce-118">В приведенном ниже примере показан вызов метода `getSelectedRange()`, который затем устанавливает желтый цвет заливки для диапазона.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="2d7ce-119">Создание книги</span><span class="sxs-lookup"><span data-stu-id="2d7ce-119">Create a workbook</span></span>

<span data-ttu-id="2d7ce-120">Ваша надстройка может создать новую книгу, отдельную от экземпляра Excel, в котором в настоящее время работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="2d7ce-121">Для этой цели в объекте Excel имеется метод `createWorkbook`.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="2d7ce-122">При вызове этого метода сразу открывается и отображается новая книга в новом экземпляре программы Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="2d7ce-123">Ваша надстройка остается открытой и запущенной в предыдущей книге.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="2d7ce-124">С помощью метода `createWorkbook` также можно создать копию существующей книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="2d7ce-125">Метод принимает в качестве необязательного параметра строковое представление XLSX-файла в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="2d7ce-126">Полученная книга будет копией этого файла, предполагая, что строковый аргумент является допустимым XLSX-файлом.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="2d7ce-127">Текущую книгу надстройки можно получить в качестве строки с кодом base64 с помощью [нарезки файлов.](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="2d7ce-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="2d7ce-128">Преобразование файла в нужную строку в кодировке base64 можно выполнить с помощью класса [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader), как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a><span data-ttu-id="2d7ce-129">Вставка копии существующей книги в текущую книгу.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-129">Insert a copy of an existing workbook into the current one</span></span>

<span data-ttu-id="2d7ce-130">В предыдущем примере показана новая книга, которая была создана из существующей книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-130">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="2d7ce-131">Вы также можете скопировать отдельные части или всю существующую книгу целиком в книгу, привязанную в настоящее время к вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-131">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="2d7ce-132">В [книге](/javascript/api/excel/excel.workbook) используется метод вставки копий таблиц целевой книги `insertWorksheetsFromBase64` в себя.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-132">A [Workbook](/javascript/api/excel/excel.workbook) has the `insertWorksheetsFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="2d7ce-133">Файл другой книги передается как строка с кодом base64, как и `Excel.createWorkbook` вызов.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-133">The other workbook's file is passed as a base64-encoded string, just like the `Excel.createWorkbook` call.</span></span> 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> <span data-ttu-id="2d7ce-134">Метод `insertWorksheetsFromBase64` поддерживается для Excel на Windows, Mac и в Интернете.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-134">The `insertWorksheetsFromBase64` method is supported for Excel on Windows, Mac, and the web.</span></span> <span data-ttu-id="2d7ce-135">Он не поддерживается для iOS.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-135">It's not supported for iOS.</span></span> <span data-ttu-id="2d7ce-136">Кроме того, Excel в Интернете этот метод не поддерживает исходные таблицы с элементами PivotTable, Chart, Comment или Slicer.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-136">Additionally, in Excel on the web this method doesn't support source worksheets with PivotTable, Chart, Comment, or Slicer elements.</span></span> <span data-ttu-id="2d7ce-137">Если эти объекты присутствуют, метод возвращает `insertWorksheetsFromBase64` `UnsupportedFeature` ошибку в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-137">If those objects are present, the `insertWorksheetsFromBase64` method returns the `UnsupportedFeature` error in Excel on the web.</span></span> 

<span data-ttu-id="2d7ce-138">В следующем примере кода показано, как вставить в текущую книгу таблицы из другой книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-138">The following code sample shows how to insert worksheets from another workbook into the current workbook.</span></span> <span data-ttu-id="2d7ce-139">Этот пример кода сначала обрабатывает файл книги с объектом и извлекает строку с кодом base64, а затем вставляет эту строку с кодом base64 в текущую [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) книгу.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-139">This code sample first processes a workbook file with a [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) object and extracts a base64-encoded string, and then it inserts this base64-encoded string into the current workbook.</span></span> <span data-ttu-id="2d7ce-140">Новые листы вставляются после листа с именем **Sheet1.**</span><span class="sxs-lookup"><span data-stu-id="2d7ce-140">The new worksheets are inserted after the worksheet named **Sheet1**.</span></span> <span data-ttu-id="2d7ce-141">Обратите внимание, что он передается в качестве параметра свойства `[]` [InsertWorksheetOptions.sheetNamesToInsert.](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert)</span><span class="sxs-lookup"><span data-stu-id="2d7ce-141">Note that `[]` is passed as the parameter for the [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) property.</span></span> <span data-ttu-id="2d7ce-142">Это означает, что все таблицы из целевой книги вставляются в текущую книгу.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-142">This means that all the worksheets from the target workbook are inserted into the current workbook.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);
            
        // Retrieve the current workbook.
        var workbook = context.workbook;
            
        // Set up the insert options. 
        var options = { 
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
            relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
        }; 
            
         // Insert the new worksheets into the current workbook.
         workbook.insertWorksheetsFromBase64(externalWorkbook, options);
         return context.sync();
    });
};

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="2d7ce-143">Защита структуры книги</span><span class="sxs-lookup"><span data-stu-id="2d7ce-143">Protect the workbook's structure</span></span>

<span data-ttu-id="2d7ce-144">Надстройка может управлять возможностью пользователя по изменению структуры книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-144">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="2d7ce-145">Свойство `protection` объекта Workbook является объектом [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) с методом `protect()`.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-145">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="2d7ce-146">В приведенном ниже примере показан основной сценарий переключения защиты структуры книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-146">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="2d7ce-147">Метод `protect` принимает необязательный строковый параметр.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-147">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="2d7ce-148">Эта строка представляет пароль, необходимый пользователю для обхода защиты и изменения структуры книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-148">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="2d7ce-149">Защиту также можно установить на уровне книги, чтобы предотвратить нежелательные изменения данных.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-149">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="2d7ce-150">Дополнительные сведения см. в разделе **Защита данных** статьи [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="2d7ce-150">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="2d7ce-151">Дополнительные сведения о защите книги в Excel см. в статье [Защита книги](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="2d7ce-151">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="2d7ce-152">Доступ к свойствам документов</span><span class="sxs-lookup"><span data-stu-id="2d7ce-152">Access document properties</span></span>

<span data-ttu-id="2d7ce-153">Объекты Workbook имеют доступ к метаданным файлов Office, называемым [свойствами документов](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="2d7ce-153">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="2d7ce-154">Свойство `properties` объекта Workbook является объектом [DocumentProperties](/javascript/api/excel/excel.documentproperties), содержащим эти значения метаданных.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-154">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="2d7ce-155">В следующем примере показано, как установить `author` свойство.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-155">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a><span data-ttu-id="2d7ce-156">Настраиваемые свойства</span><span class="sxs-lookup"><span data-stu-id="2d7ce-156">Custom properties</span></span>

<span data-ttu-id="2d7ce-157">Также можно установить настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-157">You can also define custom properties.</span></span> <span data-ttu-id="2d7ce-158">Объект DocumentProperties содержит свойство `custom`, представляющее коллекцию пар "ключ-значение" для свойств, определяемых пользователем.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-158">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="2d7ce-159">В приведенном ниже примере показано, как создать настраиваемое свойство с именем **Introduction** со значением "Hello", а затем вызвать его.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-159">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

#### <a name="worksheet-level-custom-properties"></a><span data-ttu-id="2d7ce-160">Настраиваемые свойства на уровне таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7ce-160">Worksheet-level custom properties</span></span>

<span data-ttu-id="2d7ce-161">Настраиваемые свойства также можно установить на уровне таблицы.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-161">Custom properties can also be set at the worksheet level.</span></span> <span data-ttu-id="2d7ce-162">Они похожи на настраиваемые свойства на уровне документов, за исключением того, что один и тот же ключ может повторяться в разных таблицах.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-162">These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.</span></span> <span data-ttu-id="2d7ce-163">В следующем примере показано, как создать настраиваемую свойство **WorksheetGroup** со значением "Альфа" на текущем таблице, а затем получить его.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-163">The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="2d7ce-164">Доступ к параметрам документа</span><span class="sxs-lookup"><span data-stu-id="2d7ce-164">Access document settings</span></span>

<span data-ttu-id="2d7ce-165">Параметры книги похожи на коллекцию настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-165">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="2d7ce-166">Различие заключается в том, что параметры уникальны для одного файла Excel и соединения надстройки, а свойства связаны только с файлом.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-166">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="2d7ce-167">В приведенном ниже примере показано, как создать параметр и получить к нему доступ.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-167">The following example shows how to create and access a setting.</span></span>

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

## <a name="access-application-culture-settings"></a><span data-ttu-id="2d7ce-168">Доступ к настройкам культуры приложений</span><span class="sxs-lookup"><span data-stu-id="2d7ce-168">Access application culture settings</span></span>

<span data-ttu-id="2d7ce-169">В книге есть языковые и культурные параметры, влияющие на отображение определенных данных.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-169">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="2d7ce-170">Эти параметры могут помочь локализовать данные, когда пользователи надстройки делятся книгами на разных языках и культурах.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-170">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="2d7ce-171">Ваша надстройка может использовать анализ строк для локализации формата чисел, дат и времени в зависимости от параметров культуры системы, чтобы каждый пользователь видел данные в формате своей культуры.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-171">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="2d7ce-172">`Application.cultureInfo`определяет параметры культуры системы как объект [CultureInfo.](/javascript/api/excel/excel.cultureinfo)</span><span class="sxs-lookup"><span data-stu-id="2d7ce-172">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="2d7ce-173">Это содержит параметры, такие как числовой десятичной сепаратор или формат даты.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-173">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="2d7ce-174">Некоторые параметры культуры можно [изменить с помощью Excel пользовательского интерфейса.](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)</span><span class="sxs-lookup"><span data-stu-id="2d7ce-174">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="2d7ce-175">Параметры системы сохраняются в `CultureInfo` объекте.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-175">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="2d7ce-176">Любые локальные изменения хранятся как [свойства уровня приложения,](/javascript/api/excel/excel.application)например `Application.decimalSeparator` .</span><span class="sxs-lookup"><span data-stu-id="2d7ce-176">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="2d7ce-177">В следующем примере изменяется десятичное сепараторное течение числовой строки с "," на символ, используемый в параметрах системы.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-177">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="2d7ce-178">Добавление настраиваемых XML-данных в книгу</span><span class="sxs-lookup"><span data-stu-id="2d7ce-178">Add custom XML data to the workbook</span></span>

<span data-ttu-id="2d7ce-179">Формат файла Excel Open XML **(XLSX)** позволяет надстройке внедрить настраиваемые XML-данные в книгу.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-179">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="2d7ce-180">Эти данные сохраняются с книгой независимо от надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-180">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="2d7ce-181">Книга содержит объект [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), являющийся списком объектов [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="2d7ce-181">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="2d7ce-182">Они предоставляют доступ к строкам XML и соответствующему уникальному идентификатору.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-182">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="2d7ce-183">Сохраняя эти идентификаторы как параметры, надстройка может сохранять ключи к частям XML между сеансами.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-183">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="2d7ce-184">В приведенных ниже примерах показано, как использовать настраиваемые части XML.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-184">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="2d7ce-185">В первом блоке кода показано, как внедрять XML-данные в документ.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-185">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="2d7ce-186">Выполняется сохранение списка проверяющих, а затем используются параметры книги, чтобы сохранить параметр `id` XML для будущих извлечений.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-186">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="2d7ce-187">Во втором блоке показано, как получить доступ к этим XML-данным позднее.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-187">The second block shows how to access that XML later.</span></span> <span data-ttu-id="2d7ce-188">Параметр "ContosoReviewXmlPartId" загружается и передается объекту `customXmlParts` книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-188">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="2d7ce-189">Данные XML затем печатаются в консоль.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-189">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="2d7ce-190">`CustomXMLPart.namespaceUri` заполняется только в том случае, если настраиваемый XML-элемент верхнего уровня содержит атрибут `xmlns`.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-190">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="2d7ce-191">Управление режимом вычислений</span><span class="sxs-lookup"><span data-stu-id="2d7ce-191">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="2d7ce-192">Установка режима вычислений</span><span class="sxs-lookup"><span data-stu-id="2d7ce-192">Set calculation mode</span></span>

<span data-ttu-id="2d7ce-193">По умолчанию Excel пересчитывает результаты формул при каждом изменении ячейки из ссылки.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-193">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="2d7ce-194">Производительность вашей надстройки можно улучшить путем изменения режима вычислений.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-194">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="2d7ce-195">У объекта Application есть свойство `calculationMode` типа `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-195">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="2d7ce-196">Его можно установить к следующим значениям.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-196">It can be set to the following values.</span></span>

- <span data-ttu-id="2d7ce-197">`automatic`: режим пересчета по умолчанию, при котором Excel вычисляет новые результаты формулы при каждом изменении соответствующих данных.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-197">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="2d7ce-198">`automaticExceptTables`: аналогично `automatic`, за исключением того, что игнорируются любые изменения значений таблиц.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-198">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="2d7ce-199">`manual`: вычисления выполняются только в том случае, если пользователь или надстройка запрашивает их.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-199">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="2d7ce-200">Установка типа вычислений</span><span class="sxs-lookup"><span data-stu-id="2d7ce-200">Set calculation type</span></span>

<span data-ttu-id="2d7ce-201">Объект [Application](/javascript/api/excel/excel.application) предоставляет метод применения немедленного пересчета.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-201">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="2d7ce-202">Метод `Application.calculate(calculationType)` запускает ручной пересчет с учетом указанного типа `calculationType`.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-202">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="2d7ce-203">Можно укаварить следующие значения.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-203">The following values can be specified.</span></span>

- <span data-ttu-id="2d7ce-204">`full`: пересчет всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-204">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="2d7ce-205">`fullRebuild`: проверка зависимых формул с последующим пересчетом всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-205">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="2d7ce-206">`recalculate`: пересчет формул, которые были изменены (или помечены программным путем для пересчета) с момента последнего вычисления, и зависимых от них формул во всех активных книгах.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-206">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="2d7ce-207">Дополнительные сведения о пересчете см. в статье [Изменение пересчета, итерации или точности формулы](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="2d7ce-207">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="2d7ce-208">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="2d7ce-208">Temporarily suspend calculations</span></span>

<span data-ttu-id="2d7ce-209">API Excel также позволяет надстройкам отключить вычисления до вызова `RequestContext.sync()`.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-209">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="2d7ce-210">Для этого используется `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-210">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="2d7ce-211">Используйте этот метод, если ваша надстройка изменяет большие диапазоны без необходимости доступа к данным между изменениями.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-211">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="detect-workbook-activation"></a><span data-ttu-id="2d7ce-212">Обнаружение активации книг</span><span class="sxs-lookup"><span data-stu-id="2d7ce-212">Detect workbook activation</span></span>

<span data-ttu-id="2d7ce-213">Ваша надстройка может обнаруживать при активации книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-213">Your add-in can detect when a workbook is activated.</span></span> <span data-ttu-id="2d7ce-214">Книга становится  неактивной, когда пользователь переключает фокус на другую книгу, на другое приложение или (в Excel в Интернете) на другую вкладку веб-браузера.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-214">A workbook becomes *inactive* when the user switches focus to another workbook, to another application, or (in Excel on the web) to another tab of the web browser.</span></span> <span data-ttu-id="2d7ce-215">Книга *активируется,* когда пользователь возвращает фокус в книгу.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-215">A workbook is *activated* when the user returns focus to the workbook.</span></span> <span data-ttu-id="2d7ce-216">Активация книги может вызвать функции вызова в надстройке, например освежающие данные книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-216">The workbook activation can trigger callback functions in your add-in, such as refreshing workbook data.</span></span>

<span data-ttu-id="2d7ce-217">Чтобы определить, когда книга активирована, [зарегистрируйте](excel-add-ins-events.md#register-an-event-handler) обработник событий для [события onActivated](/javascript/api/excel/excel.workbook#onActivated) книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-217">To detect when a workbook is activated, [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the [onActivated](/javascript/api/excel/excel.workbook#onActivated) event of a workbook.</span></span> <span data-ttu-id="2d7ce-218">Обработчики событий `onActivated` для события получают объект [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) при пожаре события.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-218">Event handlers for the `onActivated` event receive a [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) object when the event fires.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2d7ce-219">Событие `onActivated` не определяет, когда книга открывается.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-219">The `onActivated` event doesn't detect when a workbook is opened.</span></span> <span data-ttu-id="2d7ce-220">Это событие обнаруживает только тогда, когда пользователь переключается на уже открытую книгу.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-220">This event only detects when a user switches focus back to an already open workbook.</span></span>

<span data-ttu-id="2d7ce-221">В следующем примере кода показано, как зарегистрировать обработник событий `onActivated` и настроить функцию вызова.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-221">The following code sample shows how to register the `onActivated` event handler and set up a callback function.</span></span>

```js
Excel.run(function (context) {
    // Retrieve the workbook.
    var workbook = context.workbook;

    // Register the workbook activated event handler.
    workbook.onActivated.add(workbookActivated);

    return context.sync();
});

function workbookActivated(event) {
    Excel.run(function (context) {
        // Retrieve the workbook and load the name.
        var workbook = context.workbook;
        workbook.load("name");
        
        return context.sync().then(function () {
            // Callback function for when the workbook is activated.
            console.log(`The workbook ${workbook.name} was activated.`);
        });
    });
}
```

## <a name="save-the-workbook"></a><span data-ttu-id="2d7ce-222">Сохраните книгу.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-222">Save the workbook</span></span>

<span data-ttu-id="2d7ce-223">`Workbook.save` сохраняет книгу в постоянное хранилище.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-223">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="2d7ce-224">Метод `save` принимает один необязательный `saveBehavior` параметр, который может быть одним из следующих значений.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-224">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values.</span></span>

- <span data-ttu-id="2d7ce-225">`Excel.SaveBehavior.save` (по умолчанию): файл будет сохранен без предварительного запроса имени файла, а также место для сохранения.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-225">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="2d7ce-226">Если файл не был сохранен ранее, он будет сохранен в папке по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-226">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="2d7ce-227">Если файл уже был сохранен ранее, он будет сохранен в той же папке.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-227">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="2d7ce-228">`Excel.SaveBehavior.prompt`: если файл не был сохранен ранее, будет предложено ввести имя файла и место для сохранения.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-228">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="2d7ce-229">Если файл уже был сохранен ранее, он будет сохраняться в той же папке, и никаких дополнительных действий не потребуется.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-229">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="2d7ce-230">Если пользователь при запрос на сохранение отменяет операцию, `save` выдает исключение.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-230">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="2d7ce-231">Закрытие книги.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-231">Close the workbook</span></span>

<span data-ttu-id="2d7ce-232">`Workbook.close` закрывает книгу, а также надстройки, связанные с книгой, (приложение Excel остается открытым).</span><span class="sxs-lookup"><span data-stu-id="2d7ce-232">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="2d7ce-233">Метод `close` принимает один необязательный `closeBehavior` параметр, который может быть одним из следующих значений.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-233">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values.</span></span>

- <span data-ttu-id="2d7ce-234">`Excel.CloseBehavior.save` (по умолчанию): файл будет сохранен до закрытия.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-234">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="2d7ce-235">Если файл не был сохранен ранее, будет предложено ввести имя файла и место для сохранения.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-235">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="2d7ce-236">`Excel.CloseBehavior.skipSave`: файл будет немедленно закрыт без сохранения.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-236">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="2d7ce-237">Все несохраненные изменения будут потеряны.</span><span class="sxs-lookup"><span data-stu-id="2d7ce-237">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="2d7ce-238">См. также</span><span class="sxs-lookup"><span data-stu-id="2d7ce-238">See also</span></span>

- [<span data-ttu-id="2d7ce-239">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="2d7ce-239">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="2d7ce-240">Работа с листами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2d7ce-240">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
