---
title: Получите диапазон с помощью Excel API JavaScript
description: Узнайте, как получить диапазон с Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 3062005c1febb90749c7d129a84635f7374cd69a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744629"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>Получите диапазон с помощью Excel API JavaScript

В этой статье приводится ряд примеров получения диапазона в пределах таблицы с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемый объектом, см. в Excel[. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>Получение диапазона по адресу

Следующий пример кода получает диапазон с адресом **B2:C5** из таблицы с именем **Sample**, `address` загружает ее свойство и пишет сообщение на консоль.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    
    let range = sheet.getRange("B2:C5");
    range.load("address");
    await context.sync();
    
    console.log(`The address of the range B2:C5 is "${range.address}"`);
});
```

## <a name="get-range-by-name"></a>Получение диапазона по имени

Следующий пример кода получает диапазон, `MyRange` названный из таблицы с именем **Sample**, `address` загружает его свойство и пишет сообщение на консоль.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("MyRange");
    range.load("address");
    await context.sync();

    console.log(`The address of the range "MyRange" is "${range.address}"`);
});
```

## <a name="get-used-range"></a>Получение используемого диапазона

В следующем примере кода используется диапазон от таблицы с именем **Sample**, `address` загружается ее свойство и записывает сообщение на консоль. Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки листа, которые содержат значение или форматирование. Если весь лист пустой, `getUsedRange()` метод возвращает диапазон, состоящий только из верхнего левого элемента.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getUsedRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the used range in the worksheet is "${range.address}"`);
});
```

## <a name="get-entire-range"></a>Получение всего диапазона

Следующий пример кода получает весь диапазон таблицы из таблицы с именем **Sample**, `address` загружает ее свойство и пишет сообщение на консоль.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the entire worksheet range is "${range.address}"`);
});
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Вставьте диапазон с Excel API JavaScript](excel-add-ins-ranges-insert.md)
