---
title: Объектная модель JavaScript для Excel в надстройках Office
description: Сведения об основных типах объектов в API JavaScript для Excel и способах их использовании для создания надстроек для Excel.
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: d2972a3cc30b899340cc47c24c6792eb3e5d202c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340822"
---
# <a name="excel-javascript-object-model-in-office-add-ins"></a>Объектная модель JavaScript для Excel в надстройках Office

В этой статье описано, как создавать надстройки для Excel 2016 или более поздней версии с помощью [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md). В статье приводятся основные принципы, которые являются фундаментальными при использовании этого API, а также рекомендации по выполнению определенных задач, например чтению данных из большого диапазона или записи данных в него, изменения всех ячеек в диапазоне и т. д.

> [!IMPORTANT]
> Сведения об асинхронном типе интерфейсов API Excel и принципах их работы с книгой см. в статье [Использование модели API, зависящей от приложения](../develop/application-specific-api-model.md).  

## <a name="officejs-apis-for-excel"></a>Интерфейсы API Office.js для Excel

Надстройка Excel взаимодействует с объектами в Excel с помощью API JavaScript для Office, включающего две объектных модели JavaScript:

* **API JavaScript для Excel**. Появившийся в Office 2016 [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам.

* **Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.

Скорее всего, вы будете разрабатывать большую часть функций надстроек для Excel 2016 или более поздней версии с помощью API JavaScript для Excel, но вам также потребуются объекты из общего API. Например:

* [Context](/javascript/api/office/office.context). Объект `Context` представляет среду выполнения надстройки и предоставляет доступ к ключевым объектам API. Он состоит из данных конфигурации книги, например `contentLanguage` и `officeTheme`, а также предоставляет сведения о среде выполнения надстройки, например `host` и `platform`. Кроме того, он предоставляет метод `requirements.isSetSupported()`, с помощью которого можно проверить, поддерживается ли указанный набор обязательных элементов приложением Excel, в котором работает надстройка.
* [Document](/javascript/api/office/office.document). Объект `Document` предоставляет метод `getFileAsync()`, позволяющий скачать файл Excel, в котором работает надстройка.

На рисунке ниже показано, когда можно использовать API JavaScript для Excel или общие API.

![Различия между API JS для Excel и общими API.](../images/excel-js-api-common-api.png)

## <a name="excel-specific-object-model"></a>Объектная модель для Excel

Чтобы понять API-интерфейсы Excel, вы должны понимать, как компоненты рабочей книги связаны друг с другом.

* **Рабочая книга** содержит одну или несколько **рабочих листов**.
* **Рабочий лист** содержит коллекции тех объектов данных, которые присутствуют на отдельном листе, и предоставляет доступ к ячейкам с помощью объектов **Range**.
* **Range** представляет группу смежных клеток.
* **Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур** и других объектов визуализации данных или организации.
* **Рабочие книги** содержат коллекции некоторых из этих объектов данных (таких как **таблицы**) для всей **рабочей книги**.

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

### <a name="ranges"></a>Диапазоны

Диапазон — это группа непрерывных ячеек в рабочей книге. В надстройках обычно используется нотация в стиле A1 (например, **B3** для отдельной ячейки в столбце **B** и строке **3** или **C2:F4** для ячеек из столбцов с **C** по **F** и строк со **2-й** по **4-ю**) для определения диапазонов.

Диапазоны имеют три основных свойства: `values`, `formulas` и `format`. Эти свойства получают или устанавливают значения ячеек, формулы для оценки и визуальное форматирование ячеек.

#### <a name="range-sample"></a>Образец диапазона

В следующем примере показано, как создавать записи продаж. Эта функция использует объекты `Range` для установки значений, формул и форматов.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    let productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    await context.sync();
});
```

В этом примере создаются следующие данные в текущем листе.

![Запись о продажах, показывающая строки значений, столбец формулы и отформатированные заголовки.](../images/excel-overview-range-sample.png)

Дополнительные сведения см. в статье [Настройка и получение значений диапазона, текста или формул с помощью API JavaScript для Excel](excel-add-ins-ranges-set-get-values.md).

### <a name="charts-tables-and-other-data-objects"></a>Диаграммы, таблицы и другие объекты данных

API JavaScript для Excel могут создавать и управлять структурами данных и визуализациями в Excel. Таблицы и диаграммы являются двумя наиболее часто используемыми объектами, но API поддерживают сводные таблицы, фигуры, изображения и многое другое.

#### <a name="creating-a-table"></a>Создание таблицы

Создайте таблицы с помощью диапазонов данных. Форматирование и элементы управления таблицами (например, фильтры) автоматически применяются к диапазону.

В следующем примере создается таблица с использованием диапазонов из предыдущего примера.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    await context.sync();
});
```

Использование этого примера кода на листе с предыдущими данными создает следующую таблицу.

![Таблица сделана из предыдущего рекорда продаж.](../images/excel-overview-table-sample.png)

Дополнительные сведения см. в статье [Работа с таблицами с использованием API JavaScript для Excel](excel-add-ins-tables.md).

#### <a name="creating-a-chart"></a>Создание диаграммы

Создайте диаграммы для визуализации данных в диапазоне. API поддерживают десятки разновидностей диаграмм, каждая из которых может быть настроена в соответствии с вашими потребностями.

В следующем примере создается простая гистограмма для трех элементов, которая размещается на 100 пикселей ниже верхней части листа.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    await context.sync();
});
```

Выполнение этого примера на листе с предыдущей таблицей создает следующую диаграмму.

![Гистограмма, показывающая количества трех элементов из предыдущей записи о продажах.](../images/excel-overview-chart-sample.png)

Дополнительные сведения см. в статье [Работа с диаграммами с использованием API JavaScript для Excel](excel-add-ins-charts.md).

## <a name="see-also"></a>См. также

* [Создание первой надстройки Excel](../quickstarts/excel-quickstart-jquery.md)
* [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Оптимизация производительности API JavaScript для Excel](../excel/performance.md)
* [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
