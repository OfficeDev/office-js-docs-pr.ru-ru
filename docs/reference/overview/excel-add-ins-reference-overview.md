---
title: Обзор API JavaScript для Excel
description: ''
ms.date: 03/19/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: bf1d4642a7ceeb34eab51722a398887bb5c03fec
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871131"
---
# <a name="excel-javascript-api-overview"></a>Обзор API JavaScript для Excel

Вы можете использовать API JavaScript для Excel, чтобы создавать надстройки для Excel 2016 и более поздних версий. Ниже перечислены объекты Excel высокого уровня, доступные в API. Каждая ссылка на страницу объекта содержит описание свойств, событий и методов, доступных для объекта. Чтобы узнать больше, перейдите по соответствующим ссылкам в меню.

Для удобства ниже перечислены некоторые из основных объектов Excel. 

- [Workbook](/javascript/api/excel/excel.workbook) — объект верхнего уровня, содержащий связанные объекты книг, такие как листы, таблицы, диапазоны и т. д. Его также можно использовать для вывода списка связанных ссылок.

- [Worksheet](/javascript/api/excel/excel.worksheet). Представляет лист в книге. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): коллекция объектов **Worksheet** в книге.
    - [Worksheet Protection](/javascript/api/excel/excel.worksheetprotection): защита объекта **Worksheet**.

- [Range](/javascript/api/excel/excel.range): ячейка, строка, столбец или группа ячеек, содержащая один или несколько смежных блоков ячеек.
    - [ConditionalFormat](/javascript/api/excel/excel.conditionalformat): объект, определяющий правило и формат, которые применяются к диапазону при соблюдении условия правила.
    - [DataValidation](/javascript/api/excel/excel.datavalidation): объект, ограничивающий вводимые пользователем данные диапазоном, в основе которого лежит ряд условий.
    - [RangeSort](/javascript/api/excel/excel.rangesort): объект, управляющий операциями сортировки для диапазона.

- [Table](/javascript/api/excel/excel.table): коллекция упорядоченных ячеек для упрощения управления данными.
    - [TableCollection](/javascript/api/excel/excel.tablecollection). Коллекция таблиц в книге или на листе.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection). Коллекция всех столбцов в таблице.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection): коллекция всех строк в таблице.
    - [TableSort](/javascript/api/excel/excel.tablesort): объект, управляющий операциями сортировки в таблице.

- [Chart](/javascript/api/excel/excel.chart): объект диаграммы на листе, который является визуальным представлением базовых данных.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection): коллекция диаграмм на листе.
    
- [PivotTable](/javascript/api/excel/excel.pivottable): сводная таблица Excel, которая является иерархической группировкой и представлением данных. 
    - [PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): коллекция сводных таблиц на листе.

- [Filter](/javascript/api/excel/excel.filter): объект, управляющий фильтрацией столбца таблицы.

- [NamedItem](/javascript/api/excel/excel.nameditem): определенное имя для диапазона ячеек или значения. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection). Коллекция объектов **NamedItem** в книге.

- [Binding](/javascript/api/excel/excel.binding): абстрактный класс, представляющий привязку к разделу книги.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection): коллекция объектов **Binding** в книге.

## <a name="excel-javascript-api-open-specifications"></a>Открытая спецификация по API JavaScript для Excel

Мы публикуем новые API для надстроек Excel на странице [Открытые спецификации API](../openspec.md), чтобы вы могли делиться своим мнением. Узнайте, над какими функциями API JavaScript для Excel мы работаем, и поделитесь своим мнением о проектируемых спецификациях.

## <a name="excel-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Excel

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительны сведения о наборах обязательных элементов API JavaScript для Excel см. в статье [Наборы требований API JavaScript для Excel](../requirement-sets/excel-api-requirement-sets.md).

## <a name="excel-javascript-api-reference"></a>Справочные материалы по API JavaScript для Excel

Дополнительные сведения об API JavaScript для Excel см. в [справочной документации по API JavaScript для Excel](/javascript/api/excel).

## <a name="see-also"></a>См. также

- [Общие сведения о надстройках Excel](/office/dev/add-ins/excel/excel-add-ins-overview)
- [Обзор платформы надстроек Office](/office/dev/add-ins/overview/office-add-ins)
- [Примеры надстроек Excel на сайте GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
