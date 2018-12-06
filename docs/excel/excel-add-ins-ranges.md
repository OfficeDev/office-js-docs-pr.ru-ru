---
title: Работа с диапазонами с использованием API JavaScript для Excel
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: 4a6e0014da82956b15e11e2739f6f58fb82d5030
ms.sourcegitcommit: e2ba9d7210c921d068f40d9f689314c73ad5ab4a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/05/2018
ms.locfileid: "27156609"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a>Работа с диапазонами с использованием API JavaScript для Excel

В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для диапазонов с использованием API JavaScript для Excel. Полный список свойств и методов, поддерживаемых объектом **Range**, см. в статье [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).

## <a name="get-a-range"></a>Получение диапазона

В примерах ниже показаны различные способы получения ссылки на диапазон, расположенный на листе.

### <a name="get-range-by-address"></a>Получение диапазона по адресу

В примере кода ниже показано, как получить диапазон с адресом **B2:B5** с листа **Sample** (Пример), загрузить его свойство **address** и записать сообщение в консоль.

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

### <a name="get-range-by-name"></a>Получение диапазона по имени

В примере кода ниже показано, как получить диапазон с именем **MyRange** (Мой диапазон) с листа **Sample** (Пример), загрузить его свойство **address** и записать сообщение в консоль.

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

### <a name="get-used-range"></a>Получение используемого диапазона

В примере кода ниже показано, как получить используемый диапазон с листа **Sample** (Пример), загрузить его свойство **address** и записать сообщение в консоль. Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки листа, которые содержат значение или форматирование. Если весь лист пуст, метод **getUsedRange()** возвращает диапазон, состоящий только из левой верхней ячейки листа.

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

### <a name="get-entire-range"></a>Получение всего диапазона

В примере кода ниже показано, как получить весь диапазон листа **Sample** (Пример), загрузить его свойство **address** и записать сообщение в консоль.

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

## <a name="insert-a-range-of-cells"></a>Вставка диапазона ячеек

В примере кода ниже показано, как вставить диапазон ячеек в расположение **B4:E4** и сдвинуть другие ячейки вниз, чтобы освободить место для новых ячеек.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные перед вставкой диапазона**

![Данные в Excel перед вставкой диапазона](../images/excel-ranges-start.png)

**Данные после вставки диапазона**

![Данные в Excel после вставки диапазона](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a>Очистка диапазона ячеек

В примере кода ниже показано, как удалить все содержимое и форматирование ячеек в диапазоне **E2:E5**.  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные перед очисткой диапазона**

![Данные в Excel перед очисткой диапазона](../images/excel-ranges-start.png)

**Данные после очистки диапазона**

![Данные в Excel после очистки диапазона](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>Удаление диапазона ячеек

В примере кода ниже показано, как удалить ячейки в диапазоне **B4:E4** и сдвинуть другие ячейки вверх, чтобы заполнить место, освободившееся после удаления ячеек.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные перед удалением диапазона**

![Данные в Excel перед удалением диапазона](../images/excel-ranges-start.png)

**Данные после удаления диапазона**

![Данные в Excel после удаления диапазона](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a>Задание выделенного диапазона

В примере кода ниже показано, как выделить диапазон **B2:E6** на активном листе.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Выделенный диапазон B2:E6**

![Выделенный диапазон в Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>Получение выделенного диапазона

В примере кода ниже показано, как получить выделенный диапазон, загрузить его свойство **address** и записать сообщение в консоль. 

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-values-or-formulas"></a>Задание значений или формул

В примерах ниже показано, как задать значения и формулы для одной ячейки или диапазона ячеек.

### <a name="set-value-for-a-single-cell"></a>Задание значения для одной ячейки

В примере кода ниже показано, как присвоить ячейке **C3** значение 5, а затем настроить ширину столбцов для наилучшего размещения данных.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные перед изменением значения ячейки**

![Данные в Excel перед изменением значения ячейки](../images/excel-ranges-set-start.png)

**Данные после изменения значения ячейки**

![Данные в Excel после изменения значения ячейки](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a>Задание значений для диапазона ячеек

В примере кода ниже показано, как присвоить значения ячейкам в диапазоне **B5:D5**, а затем настроить ширину столбцов для наилучшего размещения данных.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];
    
    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные перед изменением значений ячеек**

![Данные в Excel перед изменением значений ячеек](../images/excel-ranges-set-start.png)

**Данные после изменения значений ячеек**

![Данные в Excel после изменения значений ячеек](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a>Задание формулы для одной ячейки

В примере кода ниже показано, как задать формулу для ячейки **E3**, а затем настроить ширину столбцов для наилучшего размещения данных.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные перед заданием формулы для ячейки**

![Данные в Excel перед заданием формулы для ячейки](../images/excel-ranges-start-set-formula.png)

**Данные после задания формулы для ячейки**

![Данные в Excel после задания формулы для ячейки](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a>Задание формул для диапазона ячеек

В примере кода ниже показано, как задать формулы для ячеек в диапазоне **E2:E6**, а затем настроить ширину столбцов для наилучшего размещения данных.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];
    
    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные перед заданием формул для ячеек**

![Данные в Excel перед заданием формул для ячеек](../images/excel-ranges-start-set-formula.png)

**Данные после задания формул для ячеек**

![Данные в Excel после задания формул для ячеек](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a>Получение значений, текста или формул

В примерах ниже показано, как получать значения, текст и формулы из диапазона ячеек.

### <a name="get-values-from-a-range-of-cells"></a>Получение значений из диапазона ячеек

В примере кода ниже показано, как получить диапазон **B2:E6**, загрузить его свойство **values** и записать значения из этого свойства в консоль. Свойство **values** диапазона указывает необработанные значения, содержащиеся в ячейках. Даже если некоторые ячейки в диапазоне содержат формулы, свойство **values** диапазона будет указывать необработанные значения для этих ячеек, а не формулы.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

**Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)**

![Данные в Excel после задания формул для ячеек](../images/excel-ranges-set-formulas.png)

**range.values (как записано в консоль в примере кода выше)**

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a>Получение текста из диапазона ячеек

В примере кода ниже показано, как получить диапазон **B2:E6**, загрузить его свойство **text** и записать текст из этого свойства в консоль.  Свойство **text** диапазона указывает отображаемые значения для ячеек в диапазоне. Даже если некоторые ячейки в диапазоне содержат формулы, свойство **text** диапазона будет указывать отображаемые значения для этих ячеек, а не формулы.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

**Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)**

![Данные в Excel после задания формул для ячеек](../images/excel-ranges-set-formulas.png)

**range.text (как записано в консоль в примере кода выше)**

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a>Получение формул из диапазона ячеек

В примере кода ниже показано, как получить диапазон **B2:E6**, загрузить его свойство **formulas** и записать содержимое этого свойства в консоль.  Свойство **formulas** диапазона указывает формулы для ячеек, содержащих формулы, и необработанные значения для ячеек, не содержащих формулы, в диапазоне.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

**Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)**

![Данные в Excel после задания формул для ячеек](../images/excel-ranges-set-formulas.png)

**range.formulas (как записано в консоль в примере кода выше)**

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="set-range-format"></a>Задание формата диапазона

В примерах ниже показано, как задать цвет шрифта, цвет заливки и формат чисел для ячеек в диапазоне.

### <a name="set-font-color-and-fill-color"></a>Задание цвета шрифта и цвета заливки

В примере ниже показано, как задать цвет шрифта и цвет заливки для ячеек в диапазоне **B2: E2**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные в диапазоне перед заданием цвета шрифта и цвета заливки**

![Данные в Excel перед заданием формата](../images/excel-ranges-format-before.png)

**Данные в диапазоне после задания цвета шрифта и цвета заливки**

![Данные в Excel после задания формата](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a>Задание формата чисел

В примере ниже показано, как задать формат чисел для ячеек в диапазоне **D3:E5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Данные в диапазоне перед заданием формата чисел**

![Данные в Excel перед заданием формата](../images/excel-ranges-format-font-and-fill.png)

**Данные в диапазоне после задания формата чисел**

![Данные в Excel после задания формата](../images/excel-ranges-format-numbers.png)

### <a name="conditional-formatting-of-ranges"></a>Условное форматирование диапазонов

В диапазонах может применяться форматирование к отдельным ячейкам на основе условий. Дополнительные сведения об этом см. в статье [Применение условного форматирования к диапазонам Excel](excel-add-ins-conditional-formatting.md).

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>Работа с датами с использованием подключаемого модуля Moment-MSDate

[Библиотека JavaScript Moment](https://momentjs.com/) предоставляет удобный способ использования дат и меток времени. [Подключаемый модуль Moment-MSDate](https://www.npmjs.com/package/moment-msdate) преобразует формат моментов времени в предпочитаемый для Excel. Это тот же формат, который возвращает [функция ТДАТА](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46).

В приведенном ниже коде показано, как установить для диапазона в **B4** метку момента времени.

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

Это похоже на способ получения даты из ячейки и ее преобразования в формат момента времени или другой формат, как показано в приведенном ниже коде:

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

Вашей надстройке потребуется отформатировать диапазоны, чтобы отобразить даты в более понятной для человека форме. В примере `"[$-409]m/d/yy h:mm AM/PM;@"` время отобразится как "12/3/18 3:57 PM". Дополнительные сведения о форматах чисел даты и времени см. в разделе "Рекомендации по форматам даты и времени" статьи [Рекомендации по настройке числовых форматов](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).

## <a name="copy-and-paste"></a>Копирование и вставка

> [!NOTE]
> Функция copyFrom в настоящее время доступна только в общедоступной предварительной версии (бета-версии). Для применения этой функции необходимо использовать бета-версию библиотеки в CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> Если вы используете TypeScript или ваш редактор кода использует файлы определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

Функция copyFrom диапазона реплицирует поведение копирования и вставки пользовательского интерфейса Excel. Диапазон объекта, который вызывается copyFrom, является назначением. Источник для копирования передается как диапазон или адрес строки, представляющий диапазон. В следующем примере кода копируются данные из **A1:E1** в диапазон, начиная с **G1** (который заканчивается вставкой в **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

Range.copyFrom содержит три необязательных параметра.

```ts
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
``` 

`copyType` указывает, какие данные копируются из источника в назначение. 
`“Formulas”` переносит формулы в ячейках источника и сохраняет относительное положение диапазонов этих формул. Все записи, не являющиеся формулами, копируются в исходном виде. 
`“Values”` копирует значения данных, а в случае формул — результат формулы. 
`“Formats”` копирует форматирование диапазона, включая шрифт, цвет и другие параметры форматирования, но без значений. 
`”All”` (вариант по умолчанию) копирует данные и форматирование, сохраняя формулы ячеек при их обнаружении.

`skipBlanks` устанавливает, будут ли копироваться пустые ячейки в назначение. Если значение равно true, `copyFrom` пропускает пустые ячейки в диапазоне источника. Пропущенные ячейки не перезапишут существующие данные в соответствующих им ячейках конечного диапазона. Значение по умолчанию: false.

В приведенном ниже примере кода и изображениях демонстрируется это поведение в простом сценарии. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*Прежде чем предыдущая функция была запущена.*

![Данные в Excel перед запуском метода копирования диапазона.](../images/excel-range-copyfrom-skipblanks-before.png)

*После запуска предыдущей функции.*

![Данные в Excel после запуска метода копирования диапазона.](../images/excel-range-copyfrom-skipblanks-after.png)

`transpose` определяет, переставляются ли данные в исходное расположение, то есть переключаются ли строки и столбцы. Переставленный диапазон переключается на главной диагонали, поэтому строки **1**, **2** и **3** становятся столбцами **A**, **B** и **C**. 


## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)

