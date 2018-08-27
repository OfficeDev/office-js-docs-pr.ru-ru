---
title: Работа со сводными таблицами с помощью API JavaScript для Excel
description: Использование Excel JavaScript API для создания сводных таблиц и взаимодействия с их компонентами.
ms.date: 08/17/2018
ms.openlocfilehash: aa6da2e82ab9b0c255208a86012d51db77982934
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2018
ms.locfileid: "22493980"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Работа со сводными таблицами с помощью API JavaScript для Excel

Сводные таблицы упрощают большие наборы данных. Они позволяют производить быструю манипуляцию сгруппированных данных. API JavaScript для Excel позволяет надстройке создавать сводные таблицы и взаимодействовать с их компонентами. 

Если вы не знакомы с функциональностью сводных таблиц, рекомендуем рассмотреть их в качестве конечного пользователя. См. [Создание сводной таблицы для анализа данных листа](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) как учебник для начинающих на этих инструментах. 

В этой статье приведены примеры кода для распространенных сценариев.  [Excel OpenSpec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel) предоставляет полный справочную документацию по этой функции предварительного просмотра. 

Для дальнейшего понимания API сводной таблицы см. [**PivotTable**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottable.md) и [**PivotTableCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottablecollection.md).

> [!NOTE]
> Эти примеры использования API доступны в настоящее время только в общедоступной предварительной версии (бета-версия). В этих примерах требуется предварительная версия сборки для запуска. Воспользуйтесь бета-версией библиотеки [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)или примите участие в программе[предварительная оценка Office](https://products.office.com/office-insider). Возможности сводных таблиц в настоящее время доступны в сборке 16.0.10801.20004.

## <a name="hierarchies"></a>Иерархии

Сводные таблицы организованы на основе четырёх категорий иерархии: строка, столбец, данные и фильтр. В этой статье будут использоваться следующие данные, описывающие продажи фруктов из различных ферм.

![Коллекция продажи фруктов различных типов из нескольких ферм.](../images/excel-pivots-raw-data.png)

Эти данные имеют пять иерархий: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**. Каждая иерархия может существовать только в одной из четырех категорий. Если **Type** добавлен в иерархии столбца, затем он добавляется в иерархии строки и остается только в последних.

Иерархии строк и столбцов определяют группировку данных. Например иерархии строки **Farms** будет сгруппировать все наборы данных из той же фермы. Выбор между иерархиями строк и столбцов определяет ориентацию сводных таблиц.

Иерархии данных представляют собой значения, которые нужно сгруппировать в зависимости от иерархии строк и столбцов. Сводная таблица с иерархиями строки **Farms** и данных **Crates Sold Wholesale** показывает общую сумму (по умолчанию) разных фруктов для каждой фермы.

Иерархии фильтра включают или исключаютданных изсводного документа, на основе значений в рамках отфильтрованного типа. Иерархия фильтра **Classification** с только выбранным типом**Organic** отображает данные для органических фруктов.

Вот данные фермы, вместе с сводной таблицей. Сводная таблица использует **Farm** и **Type**в качестве иерархий строк, **Crates Sold at Farm** и **Crates Sold Wholesale** - иерархий данных (с помощью статистической функции суммы по умолчанию) и **Classification** в качестве иерархий фильтра (с выбранным**Organic**). 

![Выбор данных о продажах фруктов рядом со Сводной таблицей с иерархиями строки, данных и фильтра.](../images/excel-pivot-table-and-data.png)

Эта сводная таблица может быть создана через API JavaScript или с помощью пользовательского интерфейса Excel. Оба параметра разрешают дальнейшую манипуляцию посредством надстроек.

## <a name="create-a-pivottable"></a>Создание сводной таблицы

Сводной таблице требуются имя, источник и местом назначения. Источником может быть адрес диапазона или имя таблицы (передается как `Range`, `string`, или `Table` типа). Местом назначения является адреса диапазона (представленный в виде`Range` или `string`). Следующие примеры показывают различные способы создания сводной таблицы.

### <a name="create-a-pivottable-with-range-addresses"></a>Создание сводной таблицы с помощью адресов диапазона

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Создание сводной таблицы с помощью объектов диапазона

```typescript
await Excel.run(async (context) => {    
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);
    
    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>Создание сводной таблицы на уровне рабочей книги

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Использование существующей сводной таблицы

Созданные вручную сводные таблицы, также доступны через коллекцию сводной таблицы рабочей книги или отдельных листов. 

Следующий код получает первую сводную таблицу в рабочей книге. Затем таблице дается  имя для простоты последующего использования.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Добавление строк и столбцов сводной таблицы

Строки и столбцы сводной таблицы: данные, применимые к тем значениям полей.

Добавление сводной таблицы столбца всех продаж, применимых к каждой ферме **Farm**. Добавление строк **Type** и **Classification** дополнительно разбивает данные на основе какие фрукты были проданы и была ли они органическими или нет.

![Сводная таблица со столбцом фермы и строками типа и классификации.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

Вы также можете иметь сводную таблицу только строк или столбцов.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Добавление иерархий данных сводным таблицам

Иерархии данных заполняю т сводную таблицу сведениямина основе объединения строк и столбцов. Добавление иерархий данных **Crates Sol at Farm** и **Crates Sold Wholesale** дает суммы эти цифр для каждой строки и столбца. 

В примере, обе **Farm** и **Type** являются строками, с данными продаж ящиков. 

![Сводная таблица показывает сумму всех продаж разных фруктов в ферме, в зависимости от фермы их происхождения.](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the heirarchies that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a>Изменение статистической функции

Иерархии данных имеют свои агрегированные значения. Для наборов данных чисел, это сумма по умолчанию. Свойство`summarizeBy` определяет это поведение на основе `AggregrationFunction` типа. 

В настоящее время поддерживаются следующие типы статистической функции `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, и `Automatic` (по умолчанию).

В следующих примерах кода изменяется агрегирование для средних значений данных.

```typescript
    await Excel.run(async (context) => {
        const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
        pivotTable.dataHierarchies.load("no-properties-needed");
        await context.sync();

        // changing the aggregation from the default sum to an average of all the values in the hierarchy
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        await context.sync();
    });
```

## <a name="pivottable-layouts"></a>Макеты сводной таблицы

Макет сводной таблицы определяет положение иерархий и их данных. Доступ к макету для определения диапазонов, где хранятся данные. 

На следующей диаграмме показано, какие функции макета соответствуют каким диапазонам из сводной таблицы.

![Диаграмма, показывающая, какие части сводной таблицы возвращают функции get диапазона макета.](../images/excel-pivots-layout-breakdown.png)

Приведенный ниже код показывает получение последней строки сводной таблицы данных через макет. Затем эти значения суммируются друг с другом для общего итога.


```typescript
    await Excel.run(async (context) => {
        const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
        
        // get the totals for each data hierarchy from the layout
        const range = pivotTable.layout.getDataBodyRange();
        const grandTotalRange = range.getLastRow();
        grandTotalRange.load("address");
        await context.sync();
        
        // sum the totals from the PivotTable data hierarchies and place them in a new range
        const masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
        masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
        await context.sync();
    });
```

Сводные таблицы используются три стиля макета: сжатый, структурный и табличный. Мы видели сжатый стиль в предыдущих примерах. 

В приведенных ниже примерах используется, соответственнo, структурный и табличный стили, соответственно. В примере кода показано, как переключаться между различные макетами.

### <a name="outline-layout"></a>Структурный макет

![Использование макета структуры сводной таблицы.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>Табличный макет

![Использование макета таблицы сводной таблицы.](../images/excel-pivots-tabular-layout.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    await context.sync();
    
    // cycling through layout styles
    if (pivotTable.layout.layoutType === "Compact") {
        pivotTable.layout.layoutType = "Outline";
    } else if (pivotTable.layout.layoutType === "Outline") {
        pivotTable.layout.layoutType = "Tabular";
    } else {
        pivotTable.layout.layoutType = "Compact";
    }
    
    await context.sync();
});
```

## <a name="change-hierarchy-names"></a>Изменение имен иерархий

Поля иерархии доступны для редактирования. Следующий код демонстрирует как изменить отображаемые имена двух иерархий данных.

```typescript
await Excel.run(async (context) => {
    const dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    await context.sync();
    
    // changing the displayed names of these entries
    dataHierarchies.items[0].name = "Farm Sales";
    dataHierarchies.items[1].name = "Wholesale";
    await context.sync();
});
```

## <a name="delete-a-pivottable"></a>Удаление сводной таблицы

Сводные таблицы удаляются с помощью их имени.

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

> [!NOTE]
> Мы будем рады вашим отзывам и предложениям для разработок предварительной версии. Если у вас есть комментарии, предложения или проблемы, связанные с новой API сводной таблицы, оставьте свои комментарии на репозитории[UserVoice](https://officespdev.uservoice.com/forums/224641-feature-requests-and-feedback?category_id=163563) или [repo OpenSpec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec).
