---
title: Набор обязательных элементов API JavaScript для Excel 1,7
description: Сведения о наборе требований ExcelApi 1,7
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5e923cb096c7335fbe65d18b6af0280d78be1fb2
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064860"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Новые возможности API JavaScript для Excel 1.7

Функции набора обязательных элементов API JavaScript для Excel 1.7 включают API для диаграмм, событий, рабочих листов, диапазонов, свойств документа, именованных элементов, параметров защиты и стилей.

## <a name="customize-charts"></a>Настройка диаграмм

С помощью новых API диаграмм можно создавать дополнительные типы диаграмм, добавлять ряды данных в диаграмму, задавать заголовок диаграммы, добавлять заголовок оси, добавлять отображаемые единицы, добавлять линию тренда со скользящей средней, менять линию тренда на линейную и многое другое. Вот несколько примеров:

* Ось диаграммы — получайте, задавайте, форматируйте и удаляйте единицу измерения, метку и заголовок оси на диаграмме.
* Ряды диаграммы — добавляйте, задавайте и удаляйте ряды на диаграмме.  Изменяйте маркеры рядов, порядок и размер построения.
* Линии трендов диаграммы — добавляйте, получайте и форматируйте линии тренда на диаграмме.
* Условные обозначения диаграммы — форматируйте шрифт условных обозначений на диаграмме.
* Точка диаграммы — задавайте цвет точки диаграммы.
* Подстрока заголовка диаграммы — получайте и задавайте подстроку заголовка для диаграммы.
* Тип диаграммы — параметр для создания дополнительных типов диаграмм.

## <a name="events"></a>События

API событий Excel предоставляют разнообразные обработчики событий, которые позволяют вашей надстройке автоматически запускать назначенную функцию при возникновении определенного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. Список доступных событий см. в статье [Работа с событиями с помощью API JavaScript для Excel](/office/dev/add-ins/excel/excel-add-ins-events).

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Настройка внешнего вида листов и диапазонов

С помощью новых интерфейсов API можно настроить внешний вид листов несколькими способами:

* Закрепляйте области, чтобы отображать отдельные строки или столбцы при прокрутке листа. Например, если первая строка на вашем листе содержит заголовки, вы можете закрепить эту строку, чтобы заголовки столбцов оставались видимыми при прокрутке листа.
* Изменяйте цвета вкладки листа.
* Добавляйте заголовки листов.

Внешний вид диапазонов можно настроить несколькими способами:

* Задавайте стиль ячейки для диапазона, чтобы обеспечить для всех ячеек в диапазоне единообразное форматирование. Стиль ячейки — определенный набор параметров форматирования, таких как шрифты и размеры шрифтов, форматы чисел, границы ячейки и заливка ячеек. Используйте любой из встроенных стилей ячеек Excel или создайте свой собственный стиль ячейки.
* Настройте ориентацию текста для диапазона.
* Добавляйте или изменяйте гиперссылку в диапазоне, ведущую в другое место в рабочей книге или на внешнее расположение.

## <a name="manage-document-properties"></a>Управление свойствами документа

С помощью API свойств документа можно получить доступ к встроенным свойствам документа, а также создавать и управлять настраиваемыми свойствами документа для хранения состояния книги и управления рабочим процессом и бизнес-логикой.

## <a name="copy-worksheets"></a>Копирование листов

С помощью API копирования листа вы можете копировать данные и формат с одного листа на новый рабочий лист в пределах одной книги и уменьшить объем необходимой передачи данных.

## <a name="handle-ranges-with-ease"></a>Удобная обработка диапазонов

С помощью различных API-интерфейсов диапазона можно выполнять такие действия, как получение окружающей области, получение диапазона с измененными размерами и многое другое.  Эти API позволят намного эффективнее выполнять задачи обработки и адресации диапазонов.

Дополнительно:

* Параметры защиты книги и листа — используйте эти API для защиты данных на листе и в структуре книги.
* Обновление именованного элемента — используйте этот API для обновления именованного элемента.
* Получение активной ячейки — используйте этот API для получения активной ячейки книги.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,7. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,7 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,7](/javascript/api/excel?view=excel-js-1.7)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Представляет тип диаграммы. Дополнительные сведения см. в статье Excel. ChartType.|
||[id](/javascript/api/excel/excel.chart#id)|Уникальный идентификатор диаграммы. Только для чтения.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Указывает, следует ли отображать все кнопки полей в сводной диаграмме.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[вокруг](/javascript/api/excel/excel.chartareaformat#border)|Представляет формат границы области диаграммы, включающий цвет, lineStyle и толщину. Только для чтения.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[GetItem (тип: Excel. Чартаксистипе, Group?: Excel. Чартаксисграуп)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Возвращает указанную ось, определенную по типу и группе.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[Басетимеунит](/javascript/api/excel/excel.chartaxis#basetimeunit)|Возвращает или задает базовую единицу измерений для указанной оси категории.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Возвращает или задает тип оси категории.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Представляет отображаемую единицу измерения оси. Дополнительные сведения см. в статье Excel. Чартаксисдисплайунит.|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|Представляет базу логарифма при использовании логарифмических шкал.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Представляет тип основного деления для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистиккмарк.|
||[Мажортимеунитскале](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Возвращает или задает основное значение шкалы единиц измерений для оси категории, если для свойства CategoryType установлено значение TimeScale.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Представляет тип дополнительного деления для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистиккмарк.|
||[Минортимеунитскале](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Возвращает или задает дополнительное значение шкалы единиц измерений для оси категории, если для свойства CategoryType установлено значение TimeScale.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Представляет группу для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксисграуп. Только для чтения.|
||[Кустомдисплайунит](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Представляет значение отображаемой единицы измерения настраиваемой оси.  Только для чтения. Чтобы задать это свойство, используйте метод SetCustomDisplayUnit(double).|
||[height](/javascript/api/excel/excel.chartaxis#height)|Представляет высоту оси диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Представляет расстояние от левого края оси до левой стороны области диаграммы (в пунктах).  Значение null, если ось не отображается. Только для чтения.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Представляет расстояние от верхнего края оси до верха области диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Представляет тип оси. Дополнительные сведения см. в статье Excel. Чартаксистипе.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Представляет ширину оси диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Указывает, отображает ли Microsoft Excel точки данных от последней к первой.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Представляет тип шкалы оси значений. Дополнительные сведения см. в статье Excel. Чартаксисскалетипе.|
||[Сеткатегоринамес (sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Устанавливает все имена категорий для указанной оси.|
||[Сеткустомдисплайунит (значение: число)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Задает отображаемую единицу измерения оси в виде настраиваемого значения.|
||[Шовдисплайунитлабел](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Указывает, видна ли метка отображаемой единицы измерения оси.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Представляет положение подписей делений на указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистикклабелпоситион.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Представляет количество категорий или рядов между подписями делений. Может иметь значение от 1 до 31 999 или пустую строку для автоматической настройки. Возвращаемое значение всегда является числом.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Представляет количество категорий или рядов между делениями.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Логическое значение, представляющее видимость оси.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Очищает формат границы элемента диаграммы.|
||[color](/javascript/api/excel/excel.chartborder#color)|HTML-код цвета, представляющий цвет границ в диаграмме.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Представляет тип линии границы. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Представляет толщину границы (в пунктах).|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[Элемента](/javascript/api/excel/excel.chartdatalabel#autotext)|Логическое значение, указывающее на то, генерирует ли метка данных автоматически соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах).  Значение NULL, если метка данных диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Строковое значение, представляющее код формата для метки данных.|
||[position](/javascript/api/excel/excel.chartdatalabel#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Представляет формат метки данных диаграммы.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Возвращает высоту метки данных диаграммы (в пунктах). Только для чтения. Значение NULL, если метка данных диаграммы не отображается.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Возвращает ширину метки данных диаграммы (в пунктах). Только для чтения. Значение NULL, если метка данных диаграммы не отображается.|
||[символ](/javascript/api/excel/excel.chartdatalabel#separator)|Строка, представляющая разделитель для метки данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Представляет ориентацию текста для метки данных диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах). Значение NULL, если метка данных диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чартформатстринг](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. для объекта "символы диаграммы".|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Представляет высоту условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Представляет левую (в пунктах) условные обозначения диаграммы. Значение null, если условные обозначения не отображаются.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Представляет коллекцию объектов legendEntries в условных обозначениях. Только для чтения.|
||[Шовшадов](/javascript/api/excel/excel.chartlegend#showshadow)|Указывает, имеет ли легенда тень на диаграмме.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Представляет верх условных обозначений диаграммы.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Представляет ширину условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
|[Чартлежендентри](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Представляет высоту объекта legendEntry в условных обозначениях диаграммы.|
||[индекс](/javascript/api/excel/excel.chartlegendentry#index)|Представляет индекс объекта legendEntry в условных обозначениях диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Представляет левую часть объекта legendEntry диаграммы.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Представляет верхнюю часть объекта legendEntry диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Представляет ширину объекта legendEntry в условных обозначениях диаграммы.|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Представляет видимый элемент записи условных обозначений диаграммы.|
|[Чартлежендентриколлектион](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Возвращает количество legendEntry в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Возвращает объект legendEntry по указанному индексу.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Представляет стиль линии. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Представляет толщину линии (в пунктах).|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Указывает, имеет ли точка данных метку данных. Неприменимо для поверхностных диаграмм.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Представление цветового HTML-кода для цвета фона маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Представление цветового HTML-кода для цвета переднего плана маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Представляет стиль маркера точки данных диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Возвращает метку данных точки диаграммы. Только для чтения.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[вокруг](/javascript/api/excel/excel.chartpointformat#border)|Представляет формат границы точки данных диаграммы, включающий сведения о цвете, стиле и весу. Только для чтения.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Представляет тип диаграммы для ряда. Дополнительные сведения см. в статье Excel. ChartType.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Удаляет ряд диаграммы.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Представляет размер отверстия ряда кольцевой диаграммы.  Допустимо только в doughnutExploded и кольцевых диаграммах.|
||[отсортирован](/javascript/api/excel/excel.chartseries#filtered)|Логическое значение, которое указывает, фильтруется ли ряд. Неприменимо для поверхностных диаграмм.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Представляет ширину разрывов рядов диаграммы.  Допустимо только для линейчатых диаграмм и гистограмм, а также|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Логическое значение, которое указывает, имеют ли ряды метки данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Представляет цвет фона маркеров для рядов диаграммы.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Представляет цвет переднего плана для рядов диаграммы.|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|Представляет размер маркера рядов диаграммы.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|Представляет стиль маркера рядов диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|Представляет порядок построения рядов диаграммы в группе диаграммы.|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|Представляет коллекцию линий тренда в ряду. Только для чтения.|
||[Сетбубблесизес (sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Задает размеры пузырьков для ряда диаграммы. Применяется только для пузырьковых диаграмм.|
||[setValue (sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Задает значения для ряда диаграммы.  Для точечной диаграммы это соответствует значениям оси Y.|
||[Сетксаксисвалуес (sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Задает значения оси X для ряда диаграммы.  Применяется только для точечных диаграмм.|
||[Шовшадов](/javascript/api/excel/excel.chartseries#showshadow)|Логическое значение, указывающее, есть ли у ряда теневая копия.|
||[высокое](/javascript/api/excel/excel.chartseries#smooth)|Логическое значение, которое указывает, является ли ряд плавным.  Применяется только к графикам и точечным диаграммам.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Добавить (имя?: строка, индекс?: число)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Добавляет новый ряд в коллекцию. Новый добавленный ряд не виден, пока не будут заданы значения/оси x и размеры пузырьков (в зависимости от типа диаграммы).|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[Жетсубстринг (начало: число, Length: число)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Получение подстроки заголовка диаграммы. Разрыв строки ' \n ' также подсчитывает один символ.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Представляет горизонтальное выравнивание для заголовка диаграммы.|
||[left](/javascript/api/excel/excel.charttitle#left)|Представляет расстояние от левого края заголовка диаграммы до левого края области диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается.|
||[position](/javascript/api/excel/excel.charttitle#position)|Представляет положение заголовка диаграммы. Дополнительные сведения см. в статье Excel. Чарттитлепоситион.|
||[height](/javascript/api/excel/excel.charttitle#height)|Возвращает высоту заголовка диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается. Только для чтения.|
||[width](/javascript/api/excel/excel.charttitle#width)|Возвращает ширину заголовка диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается. Только для чтения.|
||[Сетформула (формула: строка)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Задает строковое значение, представляющее формулу заголовка диаграммы с использованием нотации стиля A1.|
||[Шовшадов](/javascript/api/excel/excel.charttitle#showshadow)|Представляет логическое значение, которое определяет, имеет ли заголовок диаграммы тень.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Представляет ориентацию текста для заголовка диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.charttitle#top)|Представляет расстояние от верхнего края заголовка диаграммы до верха области диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Представляет вертикальное выравнивание для заголовка диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[вокруг](/javascript/api/excel/excel.charttitleformat#border)|Представляет формат границы заголовка диаграммы, включающий цвет, lineStyle и толщину. Только для чтения.|
|[Чарттрендлине](/javascript/api/excel/excel.charttrendline)|[Бакквардпериод](/javascript/api/excel/excel.charttrendline#backwardperiod)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Удаляет объект линии тренда.|
||[Форвардпериод](/javascript/api/excel/excel.charttrendline#forwardperiod)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[SBM](/javascript/api/excel/excel.charttrendline#intercept)|Представляет значение отсекаемого отрезка линии тренда. Можно указать в виде числового значения или пустой строки (для автоматически заданных значений). Возвращаемое значение всегда является числом.|
||[Мовингаверажепериод](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Представляет период линии тренда диаграммы. Применяется только для линии тренда с типом MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Представляет имя линии тренда. Можно указать в виде строкового значения или присвоить значение NULL для автоматических значений. Возвращаемое значение всегда является строковым|
||[Полиномиалордер](/javascript/api/excel/excel.charttrendline#polynomialorder)|Представляет порядок линии тренда диаграммы. Применяется только для линии тренда с типом полинома.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Представляет форматирование линии тренда диаграммы.|
||[Клей](/javascript/api/excel/excel.charttrendline#label)|Представляет метку линии тренда диаграммы.|
||[Шовекуатион](/javascript/api/excel/excel.charttrendline#showequation)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[Шоврскуаред](/javascript/api/excel/excel.charttrendline#showrsquared)|Значение true, если величина достоверности аппроксимации для линии тренда отображается на диаграмме.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Представляет тип линии тренда диаграммы.|
|[Чарттрендлинеколлектион](/javascript/api/excel/excel.charttrendlinecollection)|[Add (Type?: Excel. Чарттрендлинетипе)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Добавляет новую линию тренда в коллекцию линий тренда.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Возвращает количество линий тренда в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Получает объект линии тренда по индексу, который является порядком вставки в массиве элементов.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Чарттрендлинеформат](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Представляет форматирование линий диаграммы. Только для чтения.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.customproperty#key)|Возвращает ключ настраиваемого свойства. Только для чтения.|
||[type](/javascript/api/excel/excel.customproperty#type)|Получает тип значения настраиваемого свойства. Только для чтения.|
||[value](/javascript/api/excel/excel.customproperty#value)|Получает или задает значение настраиваемого свойства.|
|[Кустомпропертиколлектион](/javascript/api/excel/excel.custompropertycollection)|[Add (Key: строка, Value: Any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Создает или задает настраиваемое свойство.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Вызывается, если настраиваемое свойство не существует.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Возвращает нулевой объект, если настраиваемое свойство не существует.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Обновляет все подключения к данным в коллекции.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[Редактирование](/javascript/api/excel/excel.documentproperties#author)|Получает или задает автора книги.|
||[категории](/javascript/api/excel/excel.documentproperties#category)|Получает или задает категорию книги.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Получает или задает примечания книги.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Получает или задает компанию книги.|
||[keyword](/javascript/api/excel/excel.documentproperties#keywords)|Получает или задает ключевые слова книги.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|Получает или задает менеджера книги.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Получает дату создания книги. Только для чтения.|
||[собственный](/javascript/api/excel/excel.documentproperties#custom)|Получает коллекцию настраиваемых свойств книги. Только для чтения.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Получает последнего автора книги. Только для чтения.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Получает номер редакции книги. Только для чтения.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Получает или задает тему книги.|
||[заголовок](/javascript/api/excel/excel.documentproperties#title)|Получает или задает название книги.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Получает или задает формулу именованного элемента.  Формула всегда начинается со знака "=".|
||[Аррайвалуес](/javascript/api/excel/excel.nameditem#arrayvalues)|Возвращает объект, содержащий значения и типы именованного элемента. Только для чтения.|
|[Намедитемаррайвалуес](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Представляет типы для каждого элемента в именованном массиве элементов|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Представляет значения каждого элемента в массиве именованных элементов.|
|[Range](/javascript/api/excel/excel.range)|[Жетабсолутересизедранже (Нумровс: число, Нумколумнс: число)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Получает объект Range с той же верхней левой ячейкой, что и текущий объект Range, но с указанным количеством строк и столбцов.|
||["-изображение" ()](/javascript/api/excel/excel.range#getimage--)|Отрисовывает диапазон в виде PNG-изображения в кодировке Base64.|
||[Жетсурраундингрегион ()](/javascript/api/excel/excel.range#getsurroundingregion--)|Возвращает объект Range, представляющий область вокруг верхней левой ячейки в этом диапазоне. Это диапазон, ограниченный любым сочетанием пустых строк и столбцов, относящихся к этому диапазону.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Представляет гиперссылку для текущего диапазона.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Представляет код числового формата Excel для указанного диапазона в виде строки на языке пользователя.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Указывает, является ли текущий диапазон целым столбцом. Только для чтения.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Указывает, является ли текущий диапазон целой строкой. Только для чтения.|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|Отображает карточку для активной ячейки, если она имеет содержимое c форматированным значением.|
||[style](/javascript/api/excel/excel.range#style)|Представляет стиль текущего диапазона.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Получает или задает ориентацию текста всех ячеек в диапазоне.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Определяет, равна ли высота строки объекта Range стандартной высоте листа.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Указывает, равняется ли ширина столбца объекта Range стандартной шириной листа.|
|[Ранжехиперлинк](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Представляет целевой URL-адрес для гиперссылки.|
||[Документреференце](/javascript/api/excel/excel.rangehyperlink#documentreference)|Представляет целевую ссылку на документ для гиперссылки.|
||[Сказок](/javascript/api/excel/excel.rangehyperlink#screentip)|Представляет строку, отображаемую при наведении указателя на гиперссылку.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Представляет строку, отображаемую в верхней левой ячейке диапазона.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста в ячейке установлено на равномерное распределение.|
||[delete()](/javascript/api/excel/excel.style#delete--)|Удаляет этот стиль.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Указывает, будет ли формула скрыта, если лист защищен.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Представляет горизонтальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|Указывает, содержатся ли в стиле такие свойства, как AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel и TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|Указывает, содержатся ли в стиле такие свойства границ, как Color, ColorIndex, LineStyle и Weight.|
||[includeFont](/javascript/api/excel/excel.style#includefont)|Указывает, содержатся ли в стиле такие свойства шрифта, как Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript и Underline.|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|Указывает, содержится ли в стиле свойство NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|Указывает, содержатся ли в стиле такие внутренние свойства, как Color, ColorIndex, InvertIfNegative, Pattern, PatternColor и PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|Указывает, содержатся ли в стиле такие свойства защиты, как FormulaHidden и Locked.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа для стиля.|
||[locked](/javascript/api/excel/excel.style#locked)|Указывает, заблокирован ли объект, если лист защищен.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|Код числового формата для стиля.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|Локализованный код числового формата для стиля.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|Направление чтения для стиля.|
||[borders](/javascript/api/excel/excel.style#borders)|Коллекция Border из четырех объектов Border, представляющих стиль четырех границ.|
||[builtIn](/javascript/api/excel/excel.style#builtin)|Указывает, является ли стиль встроенным.|
||[fill](/javascript/api/excel/excel.style#fill)|Заливка стиля.|
||[font](/javascript/api/excel/excel.style#font)|Объект Font, представляющий шрифт стиля.|
||[name](/javascript/api/excel/excel.style#name)|Имя стиля.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|Ориентация текста для стиля.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Представляет вертикальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Указывает, применяет ли Microsoft Excel обтекание текстом для объекта.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Добавляет новый стиль в коллекцию.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Получает стиль по имени.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Происходит при изменении данных в ячейках в определенной таблице.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Возникает при изменении выбора в определенной таблице.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Получает адрес, представляющий измененную область таблицы на конкретном листе.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Получает тип изменения, представляющий способ запуска события Changed. Дополнительные сведения см. в статье Excel. Датачанжетипе.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область таблицы на конкретном листе.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область таблицы на конкретном листе. Может возвращать пустой объект.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Получает идентификатор таблицы, в которой изменены данные.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Возникает при изменении данных в любой таблице книги или на листе.|
|[Таблеселектиончанжедевентаргс](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Получает адрес диапазона, представляющий выбранную область таблицы на конкретном листе.|
||[Исинсидетабле](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Указывает, находится ли выделение внутри таблицы. Адрес будет бесполезным, если свойству IsInsideTable присвоено значение false.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Получает идентификатор таблицы, в которой изменено выделение.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType. Только для чтения.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменено выделение.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Жетактивецелл ()](/javascript/api/excel/excel.workbook#getactivecell--)|Получает текущую активную ячейку из книги.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Представляет все подключения к данным в книге. Только для чтения.|
||[name](/javascript/api/excel/excel.workbook#name)|Получает имя книги. Только для чтения.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Получает свойства книги. Только для чтения.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Возвращает объект защиты книги. Только для чтения.|
||[стили](/javascript/api/excel/excel.workbook#styles)|Представляет коллекцию стилей, связанных с книгой. Только для чтения.|
|[Воркбукпротектион](/javascript/api/excel/excel.workbookprotection)|[Защита (пароль?: строка)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Защищает книгу. Выдает ошибку, если книга защищена.|
||[Защита](/javascript/api/excel/excel.workbookprotection#protected)|Указывает, защищена ли книга. Только для чтения.|
||[снять защиту (пароль?: строка)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Снимает защиту с книги.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Copy (Поситионтипе?: Excel. Воркшитпоситионтипе, Релативето?: Excel. лист)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Копирует лист и размещает его в указанном положении. Возвращает скопированный лист.|
||[Жетранжебиндексес (startRow: число, startColumn: число, rowCount: число, columnCount: число)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Получает объект диапазона, начинающегося с определенных строки и столбца и занимающего определенное количество строк и столбцов.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Получает объект, который можно использовать для работы с замороженными областями на листе. Только для чтения.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Возникает при активации листа.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Происходит при изменении данных на конкретном листе.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Возникает при отключении рабочего листа.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Возникает при изменении выделенного фрагмента на определенном листе.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|Возвращает стандартную (по умолчанию) высоту всех строк на листе (в пунктах). Только для чтения.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|Возвращает или задает стандартную (по умолчанию) ширину всех столбцов на листе.|
||[Табколор](/javascript/api/excel/excel.worksheet#tabcolor)|Получает или задает цвет вкладки листа.|
|[Воркшитактиватедевентаргс](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Получает идентификатор активированного листа.|
|[Воркшитаддедевентаргс](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Получает идентификатор листа, добавленного в книгу.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Получает тип изменения, представляющий способ запуска события Changed. Дополнительные сведения см. в статье Excel. Датачанжетипе.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область конкретного листа. Может возвращать пустой объект.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Возникает при активации любого листа в книге.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Возникает при добавлении нового листа в книгу.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Возникает, когда отключается любой лист в книге.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Возникает при удалении листа из книги.|
|[Воркшитдеактиватедевентаргс](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Получает идентификатор деактивированного листа.|
|[Воркшитделетедевентаргс](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Получает идентификатор листа, удаляемого из книги.|
|[Воркшитфризепанес](/javascript/api/excel/excel.worksheetfreezepanes)|[Фризеат (Фрозенранже: строка \| Range)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Задает закрепленные ячейки в представлении активного листа.|
||[Фризеколумнс (Count?: число)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Закрепляет первый столбец (или столбцы) листа на месте.|
||[Фризеровс (Count?: число)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Закрепляет верхнюю строку (или строки) листа на месте.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Получает диапазон, описывающий закрепленные ячейки в представлении активного листа.|
||[Жетлокатионорнуллобжект ()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Получает диапазон, описывающий закрепленные ячейки в представлении активного листа.|
||[разморозить ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Удаляет все закрепленные области в листе.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[снять защиту (пароль?: строка)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Снимает защиту с листа.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[Алловедитобжектс](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Представляет параметр защиты листа, разрешающий редактирование объектов.|
||[Алловедитсценариос](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Представляет параметр защиты листа, разрешающий редактирование сценариев.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Представляет параметр защиты рабочего листа для режима выделения.|
|[Воркшитселектиончанжедевентаргс](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Получает адрес диапазона, представляющий выделенную область конкретного листа.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменено выделение.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.7)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
