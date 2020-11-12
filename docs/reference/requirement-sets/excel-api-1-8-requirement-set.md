---
title: Набор обязательных элементов API JavaScript для Excel 1,8
description: Сведения о наборе требований ExcelApi 1,8.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6454a7429276148e36431bfaffdf929a19a36d76
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996209"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Новые возможности API JavaScript для Excel 1,8

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

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,8. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,8 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,8](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[Formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Задает правый операнд, если для свойства operator задан бинарный оператор, такой как GreaterThan (левый операнд — это значение, которое пользователь пытается ввести в ячейку).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|С помощью операторов тернарного между и Нотбетвин указывает верхнюю границу операнда.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|Оператор, используемый для проверки данных.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Указывает константу перечисления Чарткатегорилабеллевел, ссылающуюся на|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Задает способ отображения пустых ячеек на диаграмме.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Определяет способ использования столбцов или строк в качестве рядов данных на диаграмме.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|True, если отображаются только видимые ячейки. False, если отображаются как видимые, так и скрытые ячейки.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Возникает при активации диаграммы.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Возникает при отключении диаграммы.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Представляет plotArea для диаграммы.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Указывает константу перечисления Чартсериеснамелевел, ссылающуюся на|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Указывает, следует ли отображать метки данных, если значение больше максимального значения на оси значений.|
||[style](/javascript/api/excel/excel.chart#style)|Задает стиль диаграммы для диаграммы.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[чартид](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Получает идентификатор активированной диаграммы.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором активирована диаграмма.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[чартид](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Получает идентификатор диаграммы, добавленной в лист.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Получает идентификатор листа, в который добавлена диаграмма.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[ориентации](/javascript/api/excel/excel.chartaxis#alignment)|Задает выравнивание для указанной Метки делений оси.|
||[исбетвинкатегориес](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Указывает, пересекают ли оси значений оси категорий между категориями.|
||[Уровневые](/javascript/api/excel/excel.chartaxis#multilevel)|Указывает, является ли ось многоуровневой.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Задает код формата для Метки делений оси.|
||[корреспондирующей](/javascript/api/excel/excel.chartaxis#offset)|Указывает расстояние между уровнями подписей и расстоянием между первым уровнем и строкой оси.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Задает заданное положение оси, в котором пересекается другая ось.|
||[поситионат](/javascript/api/excel/excel.chartaxis#positionat)|Задает указанное положение оси, в котором пересекается другая ось.|
||[Сетпоситионат (значение: число)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Задает указанную позицию оси, в которой пересекается другая ось.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Указывает угол, на который текст ориентирован на метку деления оси диаграммы.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Задает форматирование заливки диаграммы.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[Сетформула (формула: строка)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Строковое значение, представляющее формулу заголовка оси диаграммы с использованием нотации стиля A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[граница](/javascript/api/excel/excel.chartaxistitleformat#border)|Задает формат границы названия оси диаграммы, включающий цвет, lineStyle и толщину.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Задает форматирование заливки для названия оси диаграммы.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Очищает формат границы элемента диаграммы.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Возникает при активации диаграммы.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Возникает при добавлении новой диаграммы на лист.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Возникает при отключении диаграммы.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Возникает при удалении диаграммы.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[Элемента](/javascript/api/excel/excel.chartdatalabel#autotext)|Указывает, будет ли метка данных автоматически создавать соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах). |
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Строковое значение, представляющее код формата для метки данных.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Представляет формат метки данных диаграммы.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Возвращает высоту метки данных диаграммы (в пунктах).|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Возвращает ширину метки данных диаграммы (в пунктах).|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Представляет угол, на который ориентирован текст для метки данных диаграммы.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[граница](/javascript/api/excel/excel.chartdatalabelformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[Элемента](/javascript/api/excel/excel.chartdatalabels#autotext)|Указывает, должны ли метки данных автоматически создавать соответствующий текст на основе контекста.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Задает горизонтальное выравнивание для метки данных диаграммы.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Задает код формата для меток данных.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Представляет угол, к которому текст ориентирован для меток данных.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[чартид](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Получает идентификатор деактивированной диаграммы.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором деактивирована диаграмма.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[чартид](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Получает идентификатор диаграммы, удаляемой с листа.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Получает идентификатор листа, в котором удаляется диаграмма.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Задает высоту legendEntry в условных обозначениях диаграммы.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Указывает индекс legendEntry в условных обозначениях диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Указывает слева от диаграммы legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Указывает верхнюю часть диаграммы legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Представляет ширину объекта legendEntry в условных обозначениях диаграммы.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[граница](/javascript/api/excel/excel.chartlegendformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Задает значение высоты plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Задает значение insideHeight для plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Задает значение insideLeft для plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Задает значение insideTop для plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Задает значение insideWidth для plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Указывает левое значение параметра plotArea.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Указывает положение plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Задает форматирование диаграммы plotArea.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Задает верхнее значение параметра plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Задает значение Width для plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[граница](/javascript/api/excel/excel.chartplotareaformat#border)|Задает атрибуты границы для диаграммы plotArea.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Задает формат заливки объекта, включающий сведения о форматировании фона.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Задает группу для указанного ряда.|
||[развертывани](/javascript/api/excel/excel.chartseries#explosion)|Задает значение развертывания для круговой диаграммы или сектора кольцевой диаграммы.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Задает угол первого сектора круговой диаграммы или кольцевой диаграммы в градусах (по часовой стрелке вертикально).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|Значение true, если Excel инвертирует шаблон элемента, если он соответствует отрицательному числу.|
||[перекрывающееся](/javascript/api/excel/excel.chartseries#overlap)|Указывает на расположение строк и столбцов.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Представляет коллекцию всех dataLabels в ряду.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Задает размер вторичного раздела круговой диаграммы или круговой диаграммы в процентном соотношении от размера основной круговой диаграммы.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Указывает способ разделения двух разделов круговой диаграммы круговой диаграммы или круговой диаграммы.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|Значение true, если Excel назначает разные цвета или узор для каждого маркера данных.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[бакквардпериод](/javascript/api/excel/excel.charttrendline#backwardperiod)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[форвардпериод](/javascript/api/excel/excel.charttrendline#forwardperiod)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[Клей](/javascript/api/excel/excel.charttrendline#label)|Представляет метку линии тренда диаграммы.|
||[шовекуатион](/javascript/api/excel/excel.charttrendline#showequation)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[шоврскуаред](/javascript/api/excel/excel.charttrendline#showrsquared)|Значение true, если величина достоверности аппроксимации для линии тренда отображается на диаграмме.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[Элемента](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Указывает, будет ли метка линии тренда автоматически создавать соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Строковое значение, представляющее формулу подписи линии тренда диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Представляет горизонтальное выравнивание для подписи линии тренда диаграммы.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Представляет расстояние от левого края подписи линии тренда диаграммы до левого края области диаграммы (в пунктах).|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Строковое значение, представляющее код формата для подписи линии тренда.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Формат метки линии тренда диаграммы.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Возвращает высоту подписи линии тренда диаграммы (в пунктах).|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Возвращает ширину подписи линии тренда диаграммы (в пунктах).|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Строка, представляющая текст подписи линии тренда на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Представляет угол, на который ориентирован текст для подписи линии тренда диаграммы.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Представляет расстояние от верхнего края подписи линии тренда диаграммы до верха области диаграммы (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Представляет вертикальное выравнивание для подписи линии тренда диаграммы.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[граница](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Задает формат границы, включающий цвет, lineStyle и толщину.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Задает формат заливки для текущей подписи линии тренда диаграммы.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Задает атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для подписи линии тренда диаграммы.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Формула проверки настраиваемых данных.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Числовой формат DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Положение DataPivotHierarchy.|
||[поле](/javascript/api/excel/excel.datapivothierarchy#field)|Возвращает сводные поля, связанные с DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|Идентификатор DataPivotHierarchy.|
||[Сеттодефаулт ()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Сбрасывает DataPivotHierarchy до значений по умолчанию.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Указывает, следует ли отображать данные в виде определенного итогового вычисления.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Указывает, отображаются ли все элементы DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Получает DataPivotHierarchy по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Получает DataPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[Remove (DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Очищает проверку данных из текущего диапазона.|
||[ерроралерт](/javascript/api/excel/excel.datavalidation#erroralert)|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|
||[игноребланкс](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Указывает, будет ли выполняться проверка данных для пустых ячеек, по умолчанию используется значение true.|
||[сообщение](/javascript/api/excel/excel.datavalidation#prompt)|Выдавать запрос при выборе пользователем ячейки.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Тип проверки данных, подробные сведения см. в статье Excel.DataValidationType.|
||[верно](/javascript/api/excel/excel.datavalidation#valid)|Указывает, являются ли все значения ячеек допустимыми в соответствии с правилами проверки данных.|
||[правила](/javascript/api/excel/excel.datavalidation#rule)|Правило проверки данных, которое содержит различные типы условий проверки данных.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Представляет предупреждающее сообщение об ошибке.|
||[шовалерт](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Указывает, следует ли отображать диалоговое окно с сообщением об ошибке, когда пользователь вводит недопустимые данные.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Тип оповещения для проверки данных. Дополнительные сведения см. в Excel. Датавалидатионалертстиле.|
||[заголовок](/javascript/api/excel/excel.datavalidationerroralert#title)|Представляет заголовок диалогового окна предупреждения об ошибке.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Указывает сообщение приглашения.|
||[шовпромпт](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Указывает, отображается ли подсказка, когда пользователь выбирает ячейку с проверкой данных.|
||[заголовок](/javascript/api/excel/excel.datavalidationprompt#title)|Задает название приглашения.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[собственный](/javascript/api/excel/excel.datavalidationrule#custom)|Условия проверки настраиваемых данных.|
||[дата](/javascript/api/excel/excel.datavalidationrule#date)|Условия проверки данных даты.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|Условия проверки десятичных данных.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Условия проверки данных списка.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Условия проверки данных TextLength.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Условия проверки данных времени.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Условия проверки данных WholeNumber.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[Formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Задает правый операнд, если для свойства operator задан бинарный оператор, такой как GreaterThan (левый операнд — это значение, которое пользователь пытается ввести в ячейку).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|С помощью операторов тернарного между и Нотбетвин указывает верхнюю границу операнда.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|Оператор, используемый для проверки данных.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[енаблемултиплефилтеритемс](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Определяет, следует ли разрешить несколько элементов фильтра.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Положение FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Возвращает сводные поля, связанные с FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|Идентификатор FilterPivotHierarchy.|
||[Сеттодефаулт ()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Сбрасывает FilterPivotHierarchy до значений по умолчанию.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Получает FilterPivotHierarchy по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Получает FilterPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[Remove (filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Отображает или не отображает список в раскрывающемся меню ячейки, по умолчанию используется значение true.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Источник списка для проверки данных|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Имя сводного поля.|
||[id](/javascript/api/excel/excel.pivotfield#id)|Идентификатор сводного поля.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Возвращает сводные поля, связанные со сводным полем.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Определяет, следует ли отображать все элементы сводного поля.|
||[Сортбилабелс (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Сортирует сводное поле.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Промежуточные итоги сводного поля.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Получает количество полей Pivot в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Получает объект PivotField по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Получает PivotField по имени.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Имя PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Возвращает сводные поля, связанные с PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|Идентификатор PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Получает PivotHierarchy по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Получает PivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Определяет, развернут ли элемент для отображения дочерних элементов или же свернут, а дочерние элементы являются скрытыми.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Имя элемента сводной таблицы.|
||[id](/javascript/api/excel/excel.pivotitem#id)|Идентификатор элемента сводной таблицы.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Указывает, является ли PivotItem видимым.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Получает число PivotItems в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Получает объект PivotItem по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Получает PivotItem по имени.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[Жетколумнлабелранже ()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Возвращает диапазон, где находятся названия столбцов сводной таблицы.|
||[Жетдатабодиранже ()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Возвращает диапазон, где находятся значения данных сводной таблицы.|
||[Жетфилтераксисранже ()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Возвращает диапазон области фильтра сводной таблицы.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Возвращает диапазон, в котором существует сводная таблица, за исключением области фильтра.|
||[Жетровлабелранже ()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Возвращает диапазон, где находятся названия строк сводной таблицы.|
||[лайауттипе](/javascript/api/excel/excel.pivotlayout#layouttype)|Это свойство указывает PivotLayoutType всех полей в сводной таблице.|
||[шовколумнграндтоталс](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для столбцов.|
||[шовровграндтоталс](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для строк.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Это свойство указывает SubtotalLocationType всех полей в сводной таблице.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Удаляет сводную таблицу.|
||[колумнхиерарчиес](/javascript/api/excel/excel.pivottable#columnhierarchies)|Иерархии сводных столбцов сводной таблицы.|
||[Иерархии](/javascript/api/excel/excel.pivottable#datahierarchies)|Иерархии сводных данных сводной таблицы.|
||[филтерхиерарчиес](/javascript/api/excel/excel.pivottable#filterhierarchies)|Иерархии сводных фильтров сводной таблицы.|
||[иерархии](/javascript/api/excel/excel.pivottable#hierarchies)|Иерархии сводного документа сводной таблицы.|
||[макет](/javascript/api/excel/excel.pivottable#layout)|PivotLayout, описывающий макет и визуальную структуру сводной таблицы.|
||[ровхиерарчиес](/javascript/api/excel/excel.pivottable#rowhierarchies)|Иерархии сводных строк сводной таблицы.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[Add (имя: строка, источник: \| Таблица строк диапазона \| , назначение: \| строка диапазона)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Добавьте сводную таблицу на основе указанных исходных данных и вставьте ее в верхнюю левую ячейку конечного диапазона.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Возвращает объект проверки данных.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Положение RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Возвращает сводные поля, связанные с RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|Идентификатор RowColumnPivotHierarchy.|
||[Сеттодефаулт ()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Сбрасывает RowColumnPivotHierarchy до значений по умолчанию.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Получает RowColumnPivotHierarchy по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Получает RowColumnPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[Remove (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Включение событий JavaScript в текущей области задач или контентной надстройке.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|Базовое сводное поле для обоснования расчета ShowAs, если применимо на основе типа ShowAsCalculation, в противном случае значение будет пустым.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|Базовый элемент для обоснования расчета ShowAs, если применимо на основе типа ShowAsCalculation, в противном случае значение будет пустым.|
||[пересчет](/javascript/api/excel/excel.showasrule#calculation)|Расчет ShowAs для использования в сводном поле данных.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Указывает, отображается ли отступ текста автоматически, если для выравнивания текста в ячейке задано равное равномерное распределение.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|Ориентация текста для стиля.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Если для свойства Automatic установлено значение true, все остальные значения будут игнорироваться при настройке промежуточных итогов.|
||[вычисления](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[каунтнумберс](/javascript/api/excel/excel.subtotals#countnumbers)||
||[Max](/javascript/api/excel/excel.subtotals#max)||
||[минут](/javascript/api/excel/excel.subtotals#min)||
||[Продукция](/javascript/api/excel/excel.subtotals#product)||
||[стандарддевиатион](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[стандарддевиатионп](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[произведен](/javascript/api/excel/excel.subtotals#sum)||
||[различ](/javascript/api/excel/excel.subtotals#variance)||
||[варианцеп](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|Возвращает числовой идентификатор.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область таблицы на конкретном листе.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область таблицы на конкретном листе.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|Значение true, если книга открыта в режиме только для чтения.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[oncalculated](/javascript/api/excel/excel.worksheet#oncalculated)|Возникает при вычислении листа.|
||[шовгридлинес](/javascript/api/excel/excel.worksheet#showgridlines)|Указывает, видимы ли линии сетки для пользователя.|
||[шовхеадингс](/javascript/api/excel/excel.worksheet#showheadings)|Указывает, видимы ли заголовки для пользователя.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Получает идентификатор листа, в котором произошло вычисление.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[oncalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Возникает при вычислении любого листа в книге.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
