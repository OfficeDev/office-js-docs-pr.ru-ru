---
title: Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)
description: ''
ms.date: 12/14/2018
ms.openlocfilehash: 42b1127580c46120d337553fdb86a19a78b37567
ms.sourcegitcommit: 09f124fac7b2e711e1a8be562a99624627c0699e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/15/2018
ms.locfileid: "27283795"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a>Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)

Эта статья основана на сведениях из статьи [Работа с диапазонами с использованием API JavaScript для Excel (основные задачи)](excel-add-ins-ranges.md) с предоставлением примеров кода, демонстрирующих способы выполнения более сложных задач с диапазонами с использованием API JavaScript для Excel. Полный список свойств и методов, поддерживаемых объектом **Range**, см. в статье [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).

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
> Функция `Range.copyFrom` в настоящее время доступна только в общедоступной предварительной версии (бета-версии). Для применения этой функции необходимо использовать бета-версию библиотеки в CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> Если вы используете TypeScript или ваш редактор кода использует файлы определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

Функция `copyFrom` диапазона реплицирует поведение копирования и вставки пользовательского интерфейса Excel. Диапазон объекта, который вызывается `copyFrom`, является назначением.
Источник для копирования передается как диапазон или адрес строки, представляющий диапазон. В следующем примере кода копируются данные из **A1:E1** в диапазон, начиная с **G1** (который заканчивается вставкой в **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

У функции `Range.copyFrom` есть три необязательных параметра.

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` указывает, какие данные копируются из источника в назначение.

- `"Formulas"` переносит формулы в ячейках источника и сохраняет относительное положение диапазонов этих формул. Все записи, не являющиеся формулами, копируются в исходном виде.
- `"Values"` копирует значения данных, а в случае формул — результат формулы.
- `"Formats"` копирует форматирование диапазона, включая шрифт, цвет и другие параметры форматирования, но без значений.
- `"All"` (вариант по умолчанию) копирует данные и форматирование, сохраняя формулы ячеек при их обнаружении.

`skipBlanks` устанавливает, будут ли копироваться пустые ячейки в назначение. Если значение равно true, `copyFrom` пропускает пустые ячейки в диапазоне источника.
Пропущенные ячейки не перезапишут существующие данные в соответствующих им ячейках конечного диапазона. Значение по умолчанию: false.

`transpose` определяет, переставляются ли данные в исходное расположение, то есть переключаются ли строки и столбцы.
Переставленный диапазон переключается на главной диагонали, поэтому строки **1**, **2** и **3** становятся столбцами **A**, **B** и **C**.

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

![Данные в Excel перед запуском метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-before.png)

*После запуска предыдущей функции.*

![Данные в Excel после запуска метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a>Удаление дубликатов

> [!NOTE]
> Функция `removeDuplicates` объекта Range в настоящее время доступна только в общедоступной предварительной версии (бета-версии). Для применения этой функции необходимо использовать бета-версию библиотеки в CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> Если вы используете TypeScript или ваш редактор кода использует файлы определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

Функция `removeDuplicates` объекта Range удаляет строки с повторяющимися записями в указанных столбцах. Функция проверяет каждую строку в диапазоне от индекса с наименьшим значением до индекса с наибольшим значением (сверху вниз). Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне. Строки в диапазоне под удаленной строкой сдвигаются вверх. Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.

Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов. Этот массив отсчитывается от нуля относительно диапазона, а не листа. Функция также использует логический параметр, который определяет, является ли первая строка заголовком. При значении **true** верхняя строка игнорируется при поиске дубликатов. Функция `removeDuplicates` возвращает объект `RemoveDuplicatesResult`, указывающий количество удаленных строк и количество оставшихся уникальных строк.

При использовании функции `removeDuplicates` диапазона, учитывайте следующее:

- Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций. Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.
- Пустые ячейки не игнорируются функцией `removeDuplicates`. Значение пустой ячейки обрабатывается как любое другое значение. Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.

В приведенном ниже примере показано удаление записей с повторяющимися значениями в первом столбце.

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*Прежде чем предыдущая функция была запущена.*

![Данные в Excel перед запуском метода удаления дубликатов](../images/excel-ranges-remove-duplicates-before.png)

*После запуска предыдущей функции.*

![Данные в Excel после запуска метода удаления дубликатов](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>См. также

- [Работа с диапазонами с использованием API JavaScript для Excel](excel-add-ins-ranges.md)
- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)