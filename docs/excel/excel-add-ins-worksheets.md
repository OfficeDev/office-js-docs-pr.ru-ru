# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>Работа с листами с использованием API JavaScript для Excel

В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для листов с использованием API JavaScript для Excel. Полный список свойств и методов, поддерживаемых объектами **Worksheet** и **WorksheetCollection**, см. в статьях [Объект Worksheet (API JavaScript для Excel)](http://dev.office.com/reference/add-ins/excel/worksheet) и [Объект WorksheetCollection (API JavaScript для Excel)](http://dev.office.com/reference/add-ins/excel/worksheetcollection).

**Примечание.** Сведения в этой статье применимы только к обычным листам, а не к листам диаграмм или макросов.

## <a name="get-worksheets"></a>Получение листов

В примере ниже показано, как возвратить коллекцию листов, загрузить свойство **name** каждого листа и записать сообщение в консоль.

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

**Примечание.** Свойство **id** листа уникальным образом идентифицирует лист в конкретной книге, и его значение не изменяется даже при переименовании или перемещении листа. При удалении листа из книги в Excel для Mac **идентификатор** удаленного листа можно назначить новому листу (созданному после удаления).

## <a name="get-the-active-worksheet"></a>Получение активного листа

В примере кода ниже показано, как получить активный лист, загрузить его свойство **name** и записать сообщение в консоль.

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

## <a name="set-the-active-worksheet"></a>Задание активного листа

В примере кода ниже показано, как задать лист **Sample** (Пример) в качестве активного, загрузить его свойство **name** и записать сообщение в консоль. Если нет листа с таким именем, метод **activate()** создаст ошибку **ItemNotFound**.

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

## <a name="reference-worksheets-by-relative-position"></a>Ссылка на листы по их относительным положениям

В примерах ниже показано, как ссылаться на лист по его относительному положению.

### <a name="get-the-first-worksheet"></a>Получение первого листа

В примере кода ниже показано, как получить первый лист в книге, загрузить его свойство **name** и записать сообщение в консоль.

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

### <a name="get-the-last-worksheet"></a>Получение последнего листа

В примере кода ниже показано, как получить последний лист в книге, загрузить его свойство **name** и записать сообщение в консоль.

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

### <a name="get-the-next-worksheet"></a>Получение следующего листа

В примере кода ниже показано, как получить лист, следующий за активным листом, в книге, загрузить его свойство **name** и записать сообщение в консоль. Если нет листа после активного листа, метод **getNext()** создаст ошибку **ItemNotFound**.

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

### <a name="get-the-previous-worksheet"></a>Получение предыдущего листа

В примере кода ниже показано, как получить лист, предшествующий активному листу, в книге, загрузить его свойство **name** и записать сообщение в консоль. Если нет листа перед активным листом, метод **getPrevious()** создаст ошибку **ItemNotFound**.

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

## <a name="add-a-worksheet"></a>Добавление листа

В примере кода ниже показано, как добавить лист **Sample** (Пример) в рабочую книгу, загрузить его свойства **name** и **position** и записать сообщение в консоль. Новый лист будет добавлен после всех существующих листов.

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

## <a name="delete-a-worksheet"></a>Удаление листа

В примере кода ниже показано, как удалить последний лист в книге (если это не единственный лист в книге) и записать сообщение в консоль.

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

## <a name="rename-a-worksheet"></a>Переименование листа

В примере ниже показано, как изменить имя активного листа на **New Name** (Новое имя).

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a>Перемещение листа

В примере ниже показано, как переместить лист из последней позиции в книге на первую.

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

## <a name="set-worksheet-visibility"></a>Настройка видимости листа

В примерах ниже показано, как настроить видимость листа.

### <a name="hide-a-worksheet"></a>Скрытие листа

В примере кода ниже показано, как сделать лист **Sample** (Пример) скрытым, загрузить его свойство **name** и записать сообщение в консоль.

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

### <a name="unhide-a-worksheet"></a>Отмена скрытия листа

В примере кода ниже показано, как сделать лист **Sample** (Пример), загрузить его свойство **name** и записать сообщение в консоль.

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

## <a name="get-a-cell-within-a-worksheet"></a>Получение ячейки листа

В примере кода ниже показано, как получить ячейку, расположенную в строке 2 и столбце 5 листа **Sample** (Пример), загрузить его свойства **address** и **values** и записать сообщение в консоль. Значения, передаваемые в метод **getCell(row: число, column: число)**, представляют собой индексируемые с нуля номера строк и столбцов получаемой ячейки.

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

## <a name="get-a-range-within-a-worksheet"></a>Получение диапазона на листе

Примеры, в которых показано, как получить диапазон на листе, см. в статье [Работа с диапазонами с использованием API JavaScript для Excel](excel-add-ins-ranges.md).

## <a name="additional-resources"></a>Дополнительные ресурсы

- [Основные понятия API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Объект Worksheet (API JavaScript для Excel)](http://dev.office.com/reference/add-ins/excel/worksheet)
- [Объект WorksheetCollection (API JavaScript для Excel)](http://dev.office.com/reference/add-ins/excel/worksheetcollection)
