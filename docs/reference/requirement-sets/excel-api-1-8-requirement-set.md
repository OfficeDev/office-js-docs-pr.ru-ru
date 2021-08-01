---
title: Excel Набор API JavaScript 1.8
description: Сведения о наборе требований ExcelApi 1.8.
ms.date: 03/19/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 87d59bb78a00035d4dc0ff8514d3214bc93397b3
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671424"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Новые возможности в Excel API JavaScript 1.8

Функции набора обязательных элементов API JavaScript для Excel 1.8 включают API для сводных таблиц, проверку данных, диаграммы, события для диаграмм, параметры производительности и создание рабочей книги.

## <a name="pivottable"></a>Сводная таблица

Этап 2 для API сводной таблицы позволяет надстройкам устанавливать иерархии сводной таблицы. Теперь вы можете управлять данными и способом их сведения. Наша [статья о сводной таблице](../../excel/excel-add-ins-pivottables.md) содержит дополнительные сведения о новых функциональных возможностях сводной таблицы.

## <a name="data-validation"></a>Проверка данных

Проверка данных позволяет управлять данными, которые вводит в лист пользователь. Вы можете ограничить ячейки предопределенными наборами ответов или задать всплывающие предупреждения о нежелательном вводе. Узнайте больше о [добавлении проверки данных в диапазоны](../../excel/excel-add-ins-data-validation.md) уже сегодня.

## <a name="charts"></a>Диаграммы

Еще один этап выпуска API диаграмм обеспечивает дополнительный программный контроль над элементами диаграммы. Теперь у вас есть расширенный доступ к условным обозначениям, осям, линии тренда и области построения.

## <a name="events"></a>События

Для диаграмм добавлены [дополнительные](../../excel/excel-add-ins-events.md) события. Пусть ваша надстройка реагирует на взаимодействие пользователей с диаграммой. Вы также можете [включать и отключать события](../../excel/performance.md#enable-and-disable-events), запускаемые во всей книге.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.8. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.8 или более ранних, см. Excel API в наборе требований [1.8](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Указывает операнд правой руки, когда свойство оператора задано двоичному оператору, такому как GreaterThan (левая операнд — это значение, в который пользователь пытается ввести в ячейку).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|С помощью ternary operators Between and NotBetween указывается верхний операнд.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|Оператор, используемый для проверки данных.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categoryLabelLevel)|Указывает константу индексации уровня метки категорий диаграммы, ссылаясь на уровень меток исходных категорий.|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayBlanksAs)|Указывает, как пустые ячейки заданы на диаграмме.|
||[plotBy](/javascript/api/excel/excel.chart#plotBy)|Определяет способ использования столбцов или строк в качестве рядов данных на диаграмме.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotVisibleOnly)|True, если отображаются только видимые ячейки.|
||[onActivated](/javascript/api/excel/excel.chart#onActivated)|Возникает при активации диаграммы.|
||[onDeactivated](/javascript/api/excel/excel.chart#onDeactivated)|Происходит, когда диаграмма отключена.|
||[plotArea](/javascript/api/excel/excel.chart#plotArea)|Представляет область сюжета для диаграммы.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesNameLevel)|Указывает константу индексации имен на уровне серии диаграмм, ссылаясь на уровень имен исходных серий.|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showDataLabelsOverMaximum)|Указывает, следует ли показывать метки данных, если значение превышает максимальное значение оси значения.|
||[style](/javascript/api/excel/excel.chart#style)|Указывает стиль диаграммы для диаграммы.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartId)|Получает ID активированной диаграммы.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetId)|Получает ID таблицы, в которой активируется диаграмма.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartId)|Получает ID диаграммы, добавляемой в таблицу.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetId)|Получает ID таблицы, в которую добавляется диаграмма.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[выравнивание](/javascript/api/excel/excel.chartaxis#alignment)|Указывает выравнивание для указанной метки тик оси.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isBetweenCategories)|Указывает, пересекает ли ось значения ось категории между категориями.|
||[multiLevel](/javascript/api/excel/excel.chartaxis#multiLevel)|Указывает, многоуровневая ли ось.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberFormat)|Указывает код формата для метки тик оси.|
||[смещение](/javascript/api/excel/excel.chartaxis#offset)|Указывает расстояние между уровнями меток и расстоянием между первым уровнем и линией оси.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Указывает указанное положение оси, где пересекается другая ось.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionAt)|Указывает положение оси, где пересекается другая ось.|
||[setPositionAt (значение: номер)](/javascript/api/excel/excel.chartaxis#setPositionAt_value_)|Задает указанное положение оси, где пересекается другая ось.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textOrientation)|Указывает угол, на который ориентирован текст для метки тика оси диаграммы.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Указывает форматирование заполнения диаграммы.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setFormula_formula_)|Строковое значение, представляющее формулу заголовка оси диаграммы с использованием нотации стиля A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[граница](/javascript/api/excel/excel.chartaxistitleformat#border)|Указывает пограничный формат заголовка оси диаграммы, который включает цвет, листил и вес.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Указывает форматирование заполнения заголовок оси диаграммы.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear__)|Очищает формат границы элемента диаграммы.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onActivated)|Возникает при активации диаграммы.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onAdded)|Возникает при добавлении новой диаграммы в таблицу.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#onDeactivated)|Происходит, когда диаграмма отключена.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#onDeleted)|Возникает при удалении диаграммы.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autoText)|Указывает, автоматически ли метка данных создает соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalAlignment)|Представляет горизонтальное выравнивание для метки данных диаграммы.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах). |
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberFormat)|Строковое значение, представляющее код формата для метки данных.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Представляет формат метки данных диаграммы.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Возвращает высоту метки данных диаграммы (в пунктах).|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Возвращает ширину метки данных диаграммы (в пунктах).|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textOrientation)|Представляет угол, на который ориентирован текст для метки данных диаграммы.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalAlignment)|Представляет вертикальное выравнивание для метки данных диаграммы.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[граница](/javascript/api/excel/excel.chartdatalabelformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autoText)|Указывает, автоматически ли метки данных создают соответствующий текст на основе контекста.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalAlignment)|Указывает горизонтальное выравнивание для метки данных диаграммы.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberFormat)|Указывает код формата для меток данных.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textOrientation)|Представляет угол, на который ориентирован текст для меток данных.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalAlignment)|Представляет вертикальное выравнивание для метки данных диаграммы.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartId)|Получает ID отключаемой диаграммы.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetId)|Получает ID таблицы, в которой деактивируется диаграмма.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartId)|Получает ID диаграммы, удаляемой из таблицы.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetId)|Получает ID таблицы, в которой удаляется диаграмма.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Указывает высоту записи легенды в легенде диаграммы.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Указывает индекс записи легенды в легенде диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Указывает левое значение записи легенды диаграммы.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Указывает верхнюю часть записи легенды диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Представляет ширину записи легенды на диаграмме Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[граница](/javascript/api/excel/excel.chartlegendformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Указывает значение высоты области участка.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideHeight)|Указывает внутреннее значение высоты области участка.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideLeft)|Указывает внутреннее левое значение области сюжета.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insideTop)|Указывает внутреннее верхнее значение области сюжета.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insideWidth)|Указывает внутреннее значение ширины области участка.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Указывает левое значение области сюжета.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Указывает положение области сюжета.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Указывает форматирование области сюжета диаграммы.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Указывает верхнее значение области сюжета.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Указывает значение ширины области участка.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[граница](/javascript/api/excel/excel.chartplotareaformat#border)|Указывает атрибуты границы области диаграммы.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Указывает формат заполнения объекта, который включает сведения о формате фона.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisGroup)|Указывает группу для указанной серии.|
||[взрыв](/javascript/api/excel/excel.chartseries#explosion)|Указывает значение взрыва для среза круговой диаграммы или пончик-диаграммы.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstSliceAngle)|Указывает угол первого среза круговой диаграммы или пончик-диаграммы в градусах (по часовой стрелке от вертикальной).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertIfNegative)|Верно, Excel выверяет шаблон в элементе, если он соответствует отрицательному номеру.|
||[перекрытие](/javascript/api/excel/excel.chartseries#overlap)|Указывает на расположение строк и столбцов.|
||[dataLabels](/javascript/api/excel/excel.chartseries#dataLabels)|Представляет коллекцию всех меток данных в серии.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondPlotSize)|Указывает размер вторичного раздела диаграммы пирога или диаграммы с круговым пирогом в процентах от размера первичного пирога.|
||[splitType](/javascript/api/excel/excel.chartseries#splitType)|Указывает способ разделения двух разделов диаграммы "пирог-пирог" или диаграммы "планка пирога".|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varyByCategories)|True, Excel назначит каждому маркеру данных другой цвет или шаблон.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardPeriod)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardPeriod)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[метка](/javascript/api/excel/excel.charttrendline#label)|Представляет метку линии тренда диаграммы.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showEquation)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showRSquared)|Значение True, если значение r-squared для линии тренда отображается на диаграмме.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autoText)|Указывает, автоматически ли метка trendline создает соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Строковая величина, которая представляет формулу метки трендовой линии диаграммы с помощью нотации в стиле A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalAlignment)|Представляет горизонтальное выравнивание метки трендовой линии диаграммы.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Представляет расстояние в точках от левого края метки трендовой линии диаграммы до левого края области диаграммы.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberFormat)|Строковое значение, которое представляет код формата для метки trendline.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Формат метки трендовой линии диаграммы.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Возвращает высоту подписи линии тренда диаграммы (в пунктах).|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Возвращает ширину подписи линии тренда диаграммы (в пунктах).|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Строка, представляющая текст подписи линии тренда на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textOrientation)|Представляет угол, на который ориентирован текст для метки трендовой линии диаграммы.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Представляет расстояние в точках от верхнего края метки трендовой линии диаграммы до верхней части области диаграммы.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalAlignment)|Представляет вертикальное выравнивание метки трендовой линии диаграммы.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[граница](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Указывает пограничный формат, который включает цвет, литейный стил и вес.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Указывает формат заполнения текущей метки трендовой линии диаграммы.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Указывает атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для метки трендовой линии диаграммы.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Формула проверки настраиваемых данных.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberFormat)|Числовой формат DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Положение DataPivotHierarchy.|
||[поле](/javascript/api/excel/excel.datapivothierarchy#field)|Возвращает сводные поля, связанные с DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID of the DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#setToDefault__)|Сбрасывает DataPivotHierarchy до значений по умолчанию.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showAs)|Указывает, следует ли показывать данные в качестве определенного суммарного вычисления.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeBy)|Указывает, показаны ли все элементы DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add_pivotHierarchy_)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getCount__)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getItem_name_)|Получает DataPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.datapivothierarchycollection#getItemOrNullObject_name_)|Получает DataPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove_DataPivotHierarchy_)|Удаляет PivotHierarchy из текущей оси.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear__)|Очищает проверку данных из текущего диапазона.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#errorAlert)|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreBlanks)|Указывает, будет ли проверка данных выполняться на пустых ячейках.|
||[сообщение](/javascript/api/excel/excel.datavalidation#prompt)|Подсказка, когда пользователи выбирают ячейку.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Тип проверки данных см. `Excel.DataValidationType` в подробностях.|
||[допустимо](/javascript/api/excel/excel.datavalidation#valid)|Указывает, являются ли все значения ячеек допустимыми в соответствии с правилами проверки данных.|
||[правило](/javascript/api/excel/excel.datavalidation#rule)|Правило проверки данных, которое содержит различные типы критериев проверки данных.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Представляет сообщение оповещений об ошибке.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showAlert)|Указывает, следует ли показывать диалоговое окно оповещения об ошибке при вводе пользователем недействительных данных.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Тип оповещений о проверке данных см. `Excel.DataValidationAlertStyle` в подробной информации.|
||[заголовок](/javascript/api/excel/excel.datavalidationerroralert#title)|Представляет название диалоговое окно оповещений об ошибке.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Указывает сообщение запроса.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showPrompt)|Указывает, отображается ли подсказка, когда пользователь выбирает ячейку с проверкой данных.|
||[заголовок](/javascript/api/excel/excel.datavalidationprompt#title)|Указывает заголовок для запроса.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[настраиваемый](/javascript/api/excel/excel.datavalidationrule#custom)|Условия проверки настраиваемых данных.|
||[дата](/javascript/api/excel/excel.datavalidationrule#date)|Условия проверки данных даты.|
||[десятичной](/javascript/api/excel/excel.datavalidationrule#decimal)|Условия проверки десятичных данных.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Условия проверки данных списка.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textLength)|Критерии проверки данных длины текста.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Условия проверки данных времени.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholeNumber)|Все критерии проверки данных номеров.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Указывает операнд правой руки, когда свойство оператора задано двоичному оператору, такому как GreaterThan (левая операнд — это значение, в который пользователь пытается ввести в ячейку).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|С помощью ternary operators Between and NotBetween указывается верхний операнд.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|Оператор, используемый для проверки данных.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enableMultipleFilterItems)|Определяет, следует ли разрешить несколько элементов фильтра.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Положение FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Возвращает сводные поля, связанные с FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID of the FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#setToDefault__)|Сбрасывает FilterPivotHierarchy до значений по умолчанию.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add_pivotHierarchy_)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getCount__)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getItem_name_)|Получает filterPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.filterpivothierarchycollection#getItemOrNullObject_name_)|Получает FilterPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove_filterPivotHierarchy_)|Удаляет PivotHierarchy из текущей оси.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#inCellDropDown)|Указывает, следует ли отображать список в выпадаемой ячейке.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Источник списка для проверки данных|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Имя сводного поля.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Возвращает сводные поля, связанные со сводным полем.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showAllItems)|Определяет, следует ли отображать все элементы сводного поля.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortByLabels_sortBy_)|Сортирует сводное поле.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Промежуточные итоги сводного поля.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getCount__)|Получает количество поворотных полей в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getItem_name_)|Получает PivotField по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotfieldcollection#getItemOrNullObject_name_)|Получает PivotField по имени.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Имя PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Возвращает сводные поля, связанные с PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID of the PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getCount__)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getItem_name_)|Получает PivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivothierarchycollection#getItemOrNullObject_name_)|Получает PivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isExpanded)|Определяет, развернут ли элемент для отображения дочерних элементов или же свернут, а дочерние элементы являются скрытыми.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Имя элемента сводной таблицы.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Указывает, отображается ли pivotItem.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getCount__)|Получает число pivotItems в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getItem_name_)|Получает PivotItem по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotitemcollection#getItemOrNullObject_name_)|Получает PivotItem по имени.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getColumnLabelRange__)|Возвращает диапазон, где находятся названия столбцов сводной таблицы.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getDataBodyRange__)|Возвращает диапазон, где находятся значения данных сводной таблицы.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getFilterAxisRange__)|Возвращает диапазон области фильтра сводной таблицы.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getRange__)|Возвращает диапазон, в котором существует сводная таблица, за исключением области фильтра.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getRowLabelRange__)|Возвращает диапазон, где находятся названия строк сводной таблицы.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layoutType)|Это свойство указывает PivotLayoutType всех полей в сводной таблице.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showColumnGrandTotals)|Указывает, показывает ли отчет PivotTable общие итоги для столбцов.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showRowGrandTotals)|Указывает, показывает ли отчет PivotTable общие итоги для строк.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotalLocation)|Это свойство указывает все `SubtotalLocationType` поля на PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete__)|Удаляет сводную таблицу.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnHierarchies)|Иерархии сводных столбцов сводной таблицы.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#dataHierarchies)|Иерархии сводных данных сводной таблицы.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterHierarchies)|Иерархии сводных фильтров сводной таблицы.|
||[иерархии](/javascript/api/excel/excel.pivottable#hierarchies)|Иерархии сводного документа сводной таблицы.|
||[макет](/javascript/api/excel/excel.pivottable#layout)|PivotLayout, описывающий макет и визуальную структуру сводной таблицы.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowHierarchies)|Иерархии сводных строк сводной таблицы.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add_name__source__destination_)|Добавьте pivotTable на основе указанных исходных данных и вставьте его в верхней левой ячейке диапазона назначения.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#dataValidation)|Возвращает объект проверки данных.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Положение RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Возвращает сводные поля, связанные с RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID of the RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#setToDefault__)|Сбрасывает RowColumnPivotHierarchy до значений по умолчанию.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add_pivotHierarchy_)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getCount__)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getItem_name_)|Получает RowColumnPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getItemOrNullObject_name_)|Получает RowColumnPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove_rowColumnPivotHierarchy_)|Удаляет PivotHierarchy из текущей оси.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableEvents)|Добавление событий JavaScript в текущую области задач или надстройку контента.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#baseField)|PivotField на основе расчета, если применимо `ShowAs` в соответствии с `ShowAsCalculation` типом, еще `null` .|
||[baseItem](/javascript/api/excel/excel.showasrule#baseItem)|Элемент, на основе `ShowAs` расчета, если применимо в соответствии с `ShowAsCalculation` типом, еще `null` .|
||[вычисление](/javascript/api/excel/excel.showasrule#calculation)|`ShowAs`Вычисление, используемого для PivotField.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoIndent)|Указывает, будет ли текст автоматически отступным, если выравнивание текста в ячейке задано на равное распределение.|
||[textOrientation](/javascript/api/excel/excel.style#textOrientation)|Ориентация текста для стиля.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Если установлено значение , все остальные значения будут `Automatic` `true` игнорироваться при настройке `Subtotals` .|
||[среднее значение](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countNumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[продукт](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standardDeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standardDeviationP)||
||[sum](/javascript/api/excel/excel.subtotals#sum)||
||[отклонение](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#varianceP)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyId)|Возвращает числимый ID.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getRange_ctx_)|Получает диапазон, который представляет измененную область таблицы на определенном таблице.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getRangeOrNullObject_ctx_)|Получает диапазон, который представляет измененную область таблицы на определенном таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readOnly)|`true`Возвращается, если книга открыта в режиме только для чтения.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#onCalculated)|Возникает при расчете таблицы.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showGridlines)|Указывает, видны ли линии сетки пользователю.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showHeadings)|Указывает, видны ли заголовки пользователю.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetId)|Получает ID таблицы, в которой произошел расчет.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getRange_ctx_)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getRangeOrNullObject_ctx_)|Получает диапазон, представляющий измененную область конкретного листа.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#onCalculated)|Возникает при расчете любого таблицы в книге.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
