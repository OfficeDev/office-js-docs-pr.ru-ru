---
title: Набор обязательных элементов API JavaScript для Excel 1,7
description: Сведения о наборе требований ExcelApi 1,7.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ea1fe7a3d28acce2d1f4e9ff33f7b2bd31758fbd
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996237"
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

API событий Excel предоставляют разнообразные обработчики событий, которые позволяют вашей надстройке автоматически запускать назначенную функцию при возникновении определенного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. Список доступных событий см. в статье [Работа с событиями с помощью API JavaScript для Excel](../../excel/excel-add-ins-events.md).

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

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,7. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,7 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Указывает тип диаграммы.|
||[id](/javascript/api/excel/excel.chart#id)|Уникальный идентификатор диаграммы.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Указывает, следует ли отображать все кнопки полей в сводной диаграмме.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[граница](/javascript/api/excel/excel.chartareaformat#border)|Представляет формат границы области диаграммы, включающий цвет, lineStyle и толщину.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[GetItem (тип: Excel. Чартаксистипе, Group?: Excel. Чартаксисграуп)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Возвращает указанную ось, определенную по типу и группе.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[басетимеунит](/javascript/api/excel/excel.chartaxis#basetimeunit)|Указывает базовую единицу для указанной оси категорий.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Указывает тип оси категорий.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Представляет отображаемую единицу измерения оси.|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|Определяет основу логарифма при использовании логарифмических шкал.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Указывает тип основного деления для указанной оси.|
||[мажортимеунитскале](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Задает значение основной шкалы для оси категорий, если для свойства CategoryType задано значение шкалы времени.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Указывает тип вспомогательного деления для указанной оси.|
||[минортимеунитскале](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Задает значение промежуточной шкалы для оси категорий, если для свойства CategoryType задано значение шкалы времени.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Задает группу для указанной оси.|
||[кустомдисплайунит](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Задает значение отображаемой единицы измерения настраиваемой оси.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Задает высоту оси диаграммы (в пунктах).|
||[left](/javascript/api/excel/excel.chartaxis#left)|Задает расстояние от левого края оси до левого края оси (в пунктах).|
||[top](/javascript/api/excel/excel.chartaxis#top)|Задает расстояние от верхнего края оси до верха области диаграммы (в пунктах).|
||[type](/javascript/api/excel/excel.chartaxis#type)|Указывает тип оси.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Задает ширину оси диаграммы (в пунктах).|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Указывает, отображаются ли в Excel точки данных от последней к первой.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Указывает тип масштаба оси значений.|
||[Сеткатегоринамес (sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Устанавливает все имена категорий для указанной оси.|
||[Сеткустомдисплайунит (значение: число)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Задает отображаемую единицу измерения оси в виде настраиваемого значения.|
||[шовдисплайунитлабел](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Указывает, отображается ли метка отображаемой единицы оси.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Задает положение подписей делений на указанной оси.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Задает количество категорий или рядов между подписями делений.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Задает количество категорий или рядов между делениями.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Указывает, отображается ли ось.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|HTML-код цвета, представляющий цвет границ в диаграмме.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Представляет тип линии границы.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Представляет толщину границы (в пунктах).|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Значение DataLabelPosition, которое представляет положение метки данных.|
||[символ](/javascript/api/excel/excel.chartdatalabel#separator)|Строка, представляющая разделитель для метки данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Указывает, отображается ли размер пузырька метки данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Указывает, отображается ли имя категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Указывает, является ли ключ условных обозначений метки данных видимым.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Указывает, отображается ли процент меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Указывает, отображается ли имя ряда подписей данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Указывает, отображается ли значение метки данных.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Задает высоту условных обозначений на диаграмме в пунктах.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Указывает левую (в пунктах) условные обозначения на диаграмме.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Представляет коллекцию объектов legendEntries в условных обозначениях.|
||[шовшадов](/javascript/api/excel/excel.chartlegend#showshadow)|Указывает, имеет ли легенда тень на диаграмме.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Указывает верхнюю часть условных обозначений диаграммы.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Задает ширину (в пунктах) условных обозначений на диаграмме.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Представляет видимый элемент записи условных обозначений диаграммы.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Возвращает количество legendEntry в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Возвращает объект legendEntry по указанному индексу.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Представляет стиль линии.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Представляет толщину линии (в пунктах).|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Указывает, имеет ли точка данных метку данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|HTML-код цвета цвета фона маркера точки данных (например, #FF0000 представляет собой красный цвет).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|HTML-код цвета, представляющий цвет переднего плана маркера точки данных (например, #FF0000 представляет собой красный цвет).|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Представляет стиль маркера точки данных диаграммы.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Возвращает метку данных точки диаграммы.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[граница](/javascript/api/excel/excel.chartpointformat#border)|Представляет формат границы точки данных диаграммы, включающий сведения о цвете, стиле и весу.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Представляет тип диаграммы для ряда.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Удаляет ряд диаграммы.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Представляет размер отверстия ряда кольцевой диаграммы.|
||[отсортирован](/javascript/api/excel/excel.chartseries#filtered)|Указывает, фильтруется ли ряд.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Представляет ширину разрывов рядов диаграммы.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Указывает, есть ли в рядах метки данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Задает цвет фона маркеров для ряда диаграммы.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Задает цвет переднего плана маркеров ряда диаграммы.|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|Задает размер маркера для ряда диаграммы.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|Задает стиль маркера для ряда диаграммы.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|Указывает порядок построения рядов диаграммы в группе диаграммы.|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|Коллекция линий тренда в ряду.|
||[Сетбубблесизес (sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Задает размеры пузырьков для ряда диаграммы.|
||[setValue (sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Задает значения для ряда диаграммы.|
||[Сетксаксисвалуес (sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Задает значения оси X для ряда диаграммы.|
||[шовшадов](/javascript/api/excel/excel.chartseries#showshadow)|Указывает, есть ли теневая копия ряда.|
||[высокое](/javascript/api/excel/excel.chartseries#smooth)|Указывает, является ли ряд гладким.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Добавить (имя?: строка, индекс?: число)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Добавляет новый ряд в коллекцию.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[Жетсубстринг (начало: число, Length: число)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Получение подстроки заголовка диаграммы.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Задает горизонтальное выравнивание для заголовка диаграммы.|
||[left](/javascript/api/excel/excel.charttitle#left)|Задает расстояние от левого края заголовка диаграммы до левого края области диаграммы (в пунктах).|
||[position](/javascript/api/excel/excel.charttitle#position)|Представляет положение заголовка диаграммы.|
||[height](/javascript/api/excel/excel.charttitle#height)|Возвращает высоту заголовка диаграммы (в пунктах).|
||[width](/javascript/api/excel/excel.charttitle#width)|Задает ширину заголовка диаграммы (в пунктах).|
||[Сетформула (формула: строка)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Задает строковое значение, представляющее формулу заголовка диаграммы с использованием нотации стиля A1.|
||[шовшадов](/javascript/api/excel/excel.charttitle#showshadow)|Представляет логическое значение, которое определяет, имеет ли заголовок диаграммы тень.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Задает угол, на который будет ориентирован текст для заголовка диаграммы.|
||[top](/javascript/api/excel/excel.charttitle#top)|Задает расстояние от верхнего края заголовка диаграммы до верха области диаграммы (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Задает вертикальное выравнивание заголовка диаграммы.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[граница](/javascript/api/excel/excel.charttitleformat#border)|Представляет формат границы заголовка диаграммы, включающий цвет, lineStyle и толщину.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Удаляет объект линии тренда.|
||[SBM](/javascript/api/excel/excel.charttrendline#intercept)|Представляет значение отсекаемого отрезка линии тренда.|
||[мовингаверажепериод](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Представляет период линии тренда диаграммы.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Представляет имя линии тренда.|
||[полиномиалордер](/javascript/api/excel/excel.charttrendline#polynomialorder)|Представляет порядок линии тренда диаграммы.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Представляет форматирование линии тренда диаграммы.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Представляет тип линии тренда диаграммы.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[Add (Type?: Excel. Чарттрендлинетипе)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Добавляет новую линию тренда в коллекцию линий тренда.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Возвращает количество линий тренда в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Получает объект линии тренда по индексу, который является порядком вставки в массиве элементов.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Представляет форматирование линий диаграммы.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.customproperty#key)|Ключ настраиваемого свойства.|
||[type](/javascript/api/excel/excel.customproperty#type)|Тип значения, используемого для настраиваемого свойства.|
||[value](/javascript/api/excel/excel.customproperty#value)|Значение настраиваемого свойства.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[Add (Key: строка, Value: Any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Создает или задает настраиваемое свойство.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Обновляет все подключения к данным в коллекции.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[Редактирование](/javascript/api/excel/excel.documentproperties#author)|Автор книги.|
||[категории](/javascript/api/excel/excel.documentproperties#category)|Категория книги.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Комментарии к книге.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Фирма книги.|
||[keyword](/javascript/api/excel/excel.documentproperties#keywords)|Ключевые слова книги.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|Руководитель книги.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Получает дату создания книги.|
||[собственный](/javascript/api/excel/excel.documentproperties#custom)|Получает коллекцию настраиваемых свойств книги.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Получает последнего автора книги.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Получает номер редакции книги.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Тема книги.|
||[заголовок](/javascript/api/excel/excel.documentproperties#title)|Название книги.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Формула именованного элемента.|
||[аррайвалуес](/javascript/api/excel/excel.nameditem#arrayvalues)|Возвращает объект, содержащий значения и типы именованного элемента.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Представляет типы для каждого элемента в именованном массиве элементов|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Представляет значения каждого элемента в массиве именованных элементов.|
|[Range](/javascript/api/excel/excel.range)|[Жетабсолутересизедранже (Нумровс: число, Нумколумнс: число)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Получает объект Range с той же верхней левой ячейкой, что и текущий объект Range, но с указанным количеством строк и столбцов.|
||["-изображение" ()](/javascript/api/excel/excel.range#getimage--)|Отрисовывает диапазон в виде PNG-изображения в кодировке Base64.|
||[Жетсурраундингрегион ()](/javascript/api/excel/excel.range#getsurroundingregion--)|Возвращает объект Range, представляющий область вокруг верхней левой ячейки в этом диапазоне.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Представляет гиперссылку для текущего диапазона.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Представляет код числового формата Excel для заданного диапазона в соответствии с языковыми параметрами пользователя.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Указывает, является ли текущий диапазон целым столбцом.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Указывает, является ли текущий диапазон целой строкой.|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|Отображает карточку для активной ячейки, если она имеет содержимое c форматированным значением.|
||[style](/javascript/api/excel/excel.range#style)|Представляет стиль текущего диапазона.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Ориентация текста для всех ячеек в диапазоне.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Определяет, равна ли высота строки объекта Range стандартной высоте листа.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Указывает, совпадает ли ширина столбца объекта Range со стандартной шириной листа.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Представляет целевой URL-адрес для гиперссылки.|
||[документреференце](/javascript/api/excel/excel.rangehyperlink#documentreference)|Представляет целевую ссылку на документ для гиперссылки.|
||[Сказок](/javascript/api/excel/excel.rangehyperlink#screentip)|Представляет строку, отображаемую при наведении указателя на гиперссылку.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Представляет строку, отображаемую в верхней левой ячейке диапазона.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Удаляет этот стиль.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Указывает, будет ли скрыта формула при защите листа.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Представляет горизонтальное выравнивание для стиля.|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|Указывает, содержит ли стиль свойства Indent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel и TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|Указывает, включают ли стиль свойства границы цвета, ColorIndex, LineStyle и Weight.|
||[includeFont](/javascript/api/excel/excel.style#includefont)|Указывает, содержит ли стиль фон, полужирный, цвет, ColorIndex, FontStyle, курсив, имя, размер, зачеркивание, подстрочный знак, Надстрочный знак и подчеркивание шрифта.|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|Указывает, содержит ли стиль свойство NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|Указывает, включены ли в стиль внутренние свойства Color, ColorIndex, InvertIfNegative, pattern, PatternColor и PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|Указывает, содержит ли стиль свойства FormulaHidden и locked Protection.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа для стиля.|
||[locked](/javascript/api/excel/excel.style#locked)|Указывает, блокируется ли объект, когда лист защищен.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|Код числового формата для стиля.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|Локализованный код числового формата для стиля.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|Направление чтения для стиля.|
||[borders](/javascript/api/excel/excel.style#borders)|Коллекция Border из четырех объектов Border, представляющих стиль четырех границ.|
||[builtIn](/javascript/api/excel/excel.style#builtin)|Указывает, является ли этот стиль встроенным.|
||[fill](/javascript/api/excel/excel.style#fill)|Заливка стиля.|
||[font](/javascript/api/excel/excel.style#font)|Объект Font, представляющий шрифт стиля.|
||[name](/javascript/api/excel/excel.style#name)|Имя стиля.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Указывает, сжимается ли текст автоматически в соответствии с шириной доступной ширины столбца.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Задает вертикальное выравнивание для стиля.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Указывает, переносит ли Excel текст в объекте.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Добавляет новый стиль в коллекцию.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Получает стиль по имени.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Происходит при изменении данных в ячейках в определенной таблице.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Возникает при изменении выбора в определенной таблице.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Получает адрес, представляющий измененную область таблицы на конкретном листе.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Получает тип изменения, представляющий способ запуска события Changed.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Получает идентификатор таблицы, в которой изменены данные.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Возникает при изменении данных в любой таблице книги или на листе.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Получает адрес диапазона, представляющий выбранную область таблицы на конкретном листе.|
||[исинсидетабле](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Указывает, находится ли выделенный фрагмент в таблице, адрес будет бесполезен, если Исинсидетабле имеет значение false.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Получает идентификатор таблицы, в которой изменено выделение.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменено выделение.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Жетактивецелл ()](/javascript/api/excel/excel.workbook#getactivecell--)|Получает текущую активную ячейку из книги.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Представляет все подключения к данным в книге.|
||[name](/javascript/api/excel/excel.workbook#name)|Получает имя книги.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Получает свойства книги.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Возвращает объект защиты для книги.|
||[стили](/javascript/api/excel/excel.workbook#styles)|Представляет коллекцию стилей, связанных с книгой.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[Защита (пароль?: строка)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Защищает книгу.|
||[Защита](/javascript/api/excel/excel.workbookprotection#protected)|Указывает, защищена ли книга.|
||[снять защиту (пароль?: строка)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Снимает защиту с книги.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Copy (Поситионтипе?: Excel. Воркшитпоситионтипе, Релативето?: Excel. лист)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Копирует лист и размещает его в указанной позиции.|
||[Жетранжебиндексес (startRow: число, startColumn: число, rowCount: число, columnCount: число)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Получает объект диапазона, начинающегося с определенных строки и столбца и занимающего определенное количество строк и столбцов.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Получает объект, который можно использовать для работы с замороженными областями на листе.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Возникает при активации листа.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Происходит при изменении данных на конкретном листе.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Возникает при отключении рабочего листа.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Возникает при изменении выделенного фрагмента на определенном листе.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|Возвращает стандартную (по умолчанию) высоту всех строк на листе (в пунктах).|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|Задает стандартную (по умолчанию) ширину всех столбцов в рабочем листе.|
||[табколор](/javascript/api/excel/excel.worksheet#tabcolor)|Цвет вкладки листа.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Получает идентификатор активированного листа.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Получает идентификатор листа, добавленного в книгу.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Получает тип изменения, представляющий способ запуска события Changed.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Возникает при активации любого листа в книге.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Возникает при добавлении нового листа в книгу.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Возникает, когда отключается любой лист в книге.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Возникает при удалении листа из книги.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Получает идентификатор деактивированного листа.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Получает идентификатор листа, удаляемого из книги.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[Фризеат (Фрозенранже: \| строка Range)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Задает закрепленные ячейки в представлении активного листа.|
||[Фризеколумнс (Count?: число)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Закрепляет первый столбец (или столбцы) листа на месте.|
||[Фризеровс (Count?: число)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Закрепляет верхнюю строку (или строки) листа на месте.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Получает диапазон, описывающий закрепленные ячейки в представлении активного листа.|
||[Жетлокатионорнуллобжект ()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Получает диапазон, описывающий закрепленные ячейки в представлении активного листа.|
||[разморозить ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Удаляет все закрепленные области в листе.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[снять защиту (пароль?: строка)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Снимает защиту с листа.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[алловедитобжектс](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Представляет параметр защиты листа, разрешающий редактирование объектов.|
||[алловедитсценариос](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Представляет параметр защиты листа, разрешающий редактирование сценариев.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Представляет параметр защиты рабочего листа для режима выделения.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Получает адрес диапазона, представляющий выделенную область конкретного листа.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменено выделение.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
