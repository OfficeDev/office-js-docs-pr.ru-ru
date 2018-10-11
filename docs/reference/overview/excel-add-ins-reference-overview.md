# <a name="excel-javascript-api-overview"></a>Обзор API JavaScript для Excel

Вы можете использовать API JavaScript для Excel, чтобы создавать надстройки для Excel 2016 или более поздней версии. Ниже перечислены объекты Excel высокого уровня, доступные в API. Каждая ссылка на страницу объекта содержит описание свойств, связей и методов, доступных для объекта. Чтобы узнать больше, перейдите по соответствующим ссылкам в меню.

Для удобства ниже перечислены некоторые из основных объектов Excel. 

- [Workbook](/javascript/api/excel/excel.workbook) — объект верхнего уровня, содержащий связанные объекты книг, такие как листы, таблицы, диапазоны и т. д. Его также можно использовать для вывода списка связанных ссылок.

- [Worksheet](/javascript/api/excel/excel.worksheet) — представляет лист в книге. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) — коллекция объектов **Worksheet** в книге.

- [Range](/javascript/api/excel/excel.range) — представляет ячейку, строку, столбец или группу ячеек, содержащую один или несколько смежных блоков ячеек.

- [Table](/javascript/api/excel/excel.table) — представляет коллекцию упорядоченных ячеек, которая упрощает управление данными.
    - [TableCollection](/javascript/api/excel/excel.tablecollection) — коллекция таблиц в книге или на листе.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection) — коллекция всех столбцов в таблице.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection) — коллекция всех строк в таблице.

- [Chart](/javascript/api/excel/excel.chart) — представляет объект диаграммы на листе, который является визуальным представлением базовых данных.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection) — коллекция диаграмм на листе.

- [TableSort](/javascript/api/excel/excel.tablesort) — представляет объект, управляющий операциями сортировки для объектов **Table**.

- [RangeSort](/javascript/api/excel/excel.rangesort) — представляет объект, управляющий операциями сортировки для объектов **Range**.

- [Filter](/javascript/api/excel/excel.filter) — представляет объект, управляющий фильтрацией столбца таблицы.

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) — представляет защиту объекта **Worksheet**.

- [NamedItem](/javascript/api/excel/excel.nameditem) — представляет определенное имя для диапазона ячеек или значения. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection) — коллекция объектов **NamedItem** в книге.

- [Binding](/javascript/api/excel/excel.binding) — абстрактный класс, представляющий привязку к разделу книги.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection) — коллекция объектов **Binding** в книге.

## <a name="excel-javascript-api-open-specifications"></a>Открытые спецификации API JavaScript для Excel

Мы разрабатываем публикуем новые API на странице [Открытые спецификации API](../openspec.md), чтобы вы могли оставлять свои отзывы и предложения о них. Узнайте, над какими функциями для API JavaScript для Excel мы работаем, и поделитесь своим мнением о спецификациях.

## <a name="excel-javascript-api-reference"></a>Справочник по API JavaScript для Excel

Для получения подробных сведений об API JavaScript для Excel см. [Справочную документацию по API JavaScript  для Excel](/javascript/api/excel).

## <a name="see-also"></a>См. также

- [Обзор надстроек Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Обзор платформы надстроек Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Примеры надстроек Excel на сайте GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
