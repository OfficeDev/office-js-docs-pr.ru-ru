---
title: Работа с датами с Excel API JavaScript
description: Используйте подключаемый Moment-MSDate с API Excel JavaScript для работы с датами.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdfc39f12b3374d9903156b1ba71a9bbd4f296735f0ed41dac56d62243058c1d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084735"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>Работа с датами с Excel API JavaScript и Moment-MSDate плагина

В этой статье приводится пример кода, который показывает, как работать с датами с Excel API JavaScript и [плагином Moment-MSDate.](https://www.npmjs.com/package/moment-msdate) Полный список свойств и методов, поддерживаемых объектом, см. `Range` [в Excel. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>Используйте подключаемый Moment-MSDate для работы с датами

[Библиотека JavaScript Moment](https://momentjs.com/) предоставляет удобный способ использования дат и меток времени. [Подключаемый модуль Moment-MSDate](https://www.npmjs.com/package/moment-msdate) преобразует формат моментов времени в предпочитаемый для Excel. Это тот же формат, который возвращает [функция ТДАТА](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46).

В следующем коде показано, как установить диапазон **на уровне B4** до момента.

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

В следующем примере кода демонстрируется аналогичная техника, чтобы вернуть дату из ячейки и преобразовать ее в другой `Moment` формат.

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

Надстройка должна отформатирование диапазонов для отображения дат в более понятной для человека форме. Например, `"[$-409]m/d/yy h:mm AM/PM;@"` отображает "12/3/18 3:57 PM". Дополнительные сведения о форматах дат и номеров времени см. в статье "Рекомендации по датам и форматам времени" в руководстве По обзору для настройки статьи [формата](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) номеров.


## <a name="see-also"></a>См. также

- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
