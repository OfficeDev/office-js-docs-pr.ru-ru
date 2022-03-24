---
title: Работа с датами с Excel API JavaScript
description: Используйте подключаемый Moment-MSDate с API Excel JavaScript для работы с датами.
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7ca6e0eacab7aab0308b2e397f313a8e07b59777
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745075"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>Работа с датами с Excel API JavaScript и Moment-MSDate плагина

В этой статье приводится пример кода, который показывает, как работать с датами с Excel API JavaScript и [плагином Moment-MSDate](https://www.npmjs.com/package/moment-msdate). Полный список свойств `Range` и методов, поддерживаемых объектом, см. [в Excel. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>Использование подключаемого Moment-MSDate для работы с датами

[Библиотека JavaScript Moment](https://momentjs.com/) предоставляет удобный способ использования дат и меток времени. [Подключаемый модуль Moment-MSDate](https://www.npmjs.com/package/moment-msdate) преобразует формат моментов времени в предпочитаемый для Excel. Это тот же формат, который возвращает [функция ТДАТА](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46).

В следующем коде показано, как установить диапазон **на уровне B4** до момента.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let now = Date.now();
    let nowMoment = moment(now);
    let nowMS = nowMoment.toOADate();

    let dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    await context.sync();
});
```

В следующем примере кода демонстрируется аналогичная техника, чтобы вернуть дату из ячейки и преобразовать ее в другой `Moment` формат.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let dateRange = sheet.getRange("B4");
    dateRange.load("values");

    await context.sync();

    let nowMS = dateRange.values[0][0];

    // Log the date as a moment.
    let nowMoment = moment.fromOADate(nowMS);
    console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

    // Log the date as a UNIX-style timestamp.
    let now = nowMoment.unix();
    console.log(`get (timestamp): ${now}`);
});
```

Надстройка должна отформатирование диапазонов для отображения дат в более понятной для человека форме. Например, отображает `"[$-409]m/d/yy h:mm AM/PM;@"` "12/3/18 3:57 PM". Дополнительные сведения о форматах дат и номеров времени см. в статье "Рекомендации по датам и форматам времени" в руководстве По обзору для настройки статьи [формата](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) номеров.

## <a name="see-also"></a>См. также

- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
