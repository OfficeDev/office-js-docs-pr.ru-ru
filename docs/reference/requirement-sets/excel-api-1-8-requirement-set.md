---
title: Набор обязательных элементов API JavaScript для Excel 1,8
description: Сведения о наборе требований ExcelApi 1,8
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a5adcf56654070ca2a8336385f73062c34e90e1d
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772011"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Новые возможности API JavaScript для Excel 1.8

Функции набора обязательных элементов API JavaScript для Excel 1.8 включают API для сводных таблиц, проверку данных, диаграммы, события для диаграмм, параметры производительности и создание рабочей книги.

## <a name="pivottable"></a>Сводная таблица

Этап 2 для API сводной таблицы позволяет надстройкам устанавливать иерархии сводной таблицы. Теперь вы можете управлять данными и способом их сведения. Наша [статья о сводной таблице](/office/dev/add-ins/excel/excel-add-ins-pivottables) содержит дополнительные сведения о новых функциональных возможностях сводной таблицы.

## <a name="data-validation"></a>Проверка данных

Проверка данных позволяет управлять данными, которые вводит в лист пользователь. Вы можете ограничить ячейки предопределенными наборами ответов или задать всплывающие предупреждения о нежелательном вводе. Узнайте больше о [добавлении проверки данных в диапазоны](/office/dev/add-ins/excel/excel-add-ins-data-validation) уже сегодня.

## <a name="charts"></a>Диаграммы

Еще один этап выпуска API диаграмм обеспечивает дополнительный программный контроль над элементами диаграммы. Теперь у вас есть расширенный доступ к условным обозначениям, осям, линии тренда и области построения.

## <a name="events"></a>События

Для диаграмм добавлены [дополнительные](/office/dev/add-ins/excel/excel-add-ins-events) события. Пусть ваша надстройка реагирует на взаимодействие пользователей с диаграммой. Вы также можете [включать и отключать события](/office/dev/add-ins/excel/performance#enable-and-disable-events), запускаемые во всей книге.

## <a name="api-list"></a>Список API

| Класс | Поля | Описание |
|:---|:---|:---|
|[Басикдатавалидатион](/javascript/api/excel/excel.basicdatavalidation)|[Formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Задает правый операнд, если для свойства operator задан бинарный оператор, такой как GreaterThan (левый операнд — это значение, которое пользователь пытается ввести в ячейку). С помощью операторов тернарного между и Нотбетвин задает нижнюю границу операнда.|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|С помощью операторов тернарного между и Нотбетвин указывает верхнюю границу операнда. Не используется с двоичными операторами, например GreaterThan.|
||[or](/javascript/api/excel/excel.basicdatavalidation#operator)|Оператор, используемый для проверки данных.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Возвращает или задает константу перечисления Чарткатегорилабеллевел, ссылающуюся на|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Возвращает или задает способ отображения пустых ячеек на диаграмме. Для чтения и записи.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Возвращает или задает способ использования столбцов или строк в качестве рядов данных на диаграмме. Для чтения и записи.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|True, если отображаются только видимые ячейки.False, если отображаются как видимые, так и скрытые ячейки. Для чтения и записи.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Возникает при активации диаграммы.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Возникает при отключении диаграммы.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Представляет plotArea для диаграммы.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Возвращает или задает константу перечисления Чартсериеснамелевел, ссылающуюся на|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Представляет, нужно ли отображать метки данных, если значение больше максимального на оси значений.|
||[style](/javascript/api/excel/excel.chart#style)|Возвращает или задает стиль для диаграммы. Для чтения и записи.|
|[Чартактиватедевентаргс](/javascript/api/excel/excel.chartactivatedeventargs)|[Чартид](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Получает идентификатор активированной диаграммы.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором активирована диаграмма.|
|[Чартаддедевентаргс](/javascript/api/excel/excel.chartaddedeventargs)|[Чартид](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Получает идентификатор диаграммы, добавленной в лист.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Получает идентификатор листа, в который добавлена диаграмма.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[ориентации](/javascript/api/excel/excel.chartaxis#alignment)|Представляет выравнивание для указанной метки делений оси. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[Исбетвинкатегориес](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Указывает, пересекает ли ось значений ось категорий между категориями.|
||[Уровневые](/javascript/api/excel/excel.chartaxis#multilevel)|Указывает, является ли ось многоуровневой или нет.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Представляет код формата для метки делений оси.|
||[корреспондирующей](/javascript/api/excel/excel.chartaxis#offset)|Представляет расстояние между уровнями меток и расстояние между первым уровнем и линией оси. Значение должно быть целым числом от 0 до 1000.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Представляет указанное положение оси в месте, где ее пересекает другая ось. Дополнительные сведения см. в статье Excel. Чартаксиспоситион.|
||[Поситионат](/javascript/api/excel/excel.chartaxis#positionat)|Представляет указанное положение оси в месте, где ее пересекает другая ось. Чтобы задать это свойство, следует использовать метод SetPositionAt(double).|
||[Сетпоситионат (значение: число)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Задает указанное положение оси в месте, где ее пересекает другая ось.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Представляет ориентацию текста для метки делений оси. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
|[Чартаксисдата](/javascript/api/excel/excel.chartaxisdata)|[ориентации](/javascript/api/excel/excel.chartaxisdata#alignment)|Представляет выравнивание для указанной метки делений оси. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[Исбетвинкатегориес](/javascript/api/excel/excel.chartaxisdata#isbetweencategories)|Указывает, пересекает ли ось значений ось категорий между категориями.|
||[Уровневые](/javascript/api/excel/excel.chartaxisdata#multilevel)|Указывает, является ли ось многоуровневой или нет.|
||[numberFormat](/javascript/api/excel/excel.chartaxisdata#numberformat)|Представляет код формата для метки делений оси.|
||[корреспондирующей](/javascript/api/excel/excel.chartaxisdata#offset)|Представляет расстояние между уровнями меток и расстояние между первым уровнем и линией оси. Значение должно быть целым числом от 0 до 1000.|
||[position](/javascript/api/excel/excel.chartaxisdata#position)|Представляет указанное положение оси в месте, где ее пересекает другая ось. Дополнительные сведения см. в статье Excel. Чартаксиспоситион.|
||[Поситионат](/javascript/api/excel/excel.chartaxisdata#positionat)|Представляет указанное положение оси в месте, где ее пересекает другая ось. Чтобы задать это свойство, следует использовать метод SetPositionAt(double).|
||[textOrientation](/javascript/api/excel/excel.chartaxisdata#textorientation)|Представляет ориентацию текста для метки делений оси. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Представляет форматирование заливки диаграммы. Только для чтения.|
|[Чартаксислоадоптионс](/javascript/api/excel/excel.chartaxisloadoptions)|[ориентации](/javascript/api/excel/excel.chartaxisloadoptions#alignment)|Представляет выравнивание для указанной метки делений оси. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[Исбетвинкатегориес](/javascript/api/excel/excel.chartaxisloadoptions#isbetweencategories)|Указывает, пересекает ли ось значений ось категорий между категориями.|
||[Уровневые](/javascript/api/excel/excel.chartaxisloadoptions#multilevel)|Указывает, является ли ось многоуровневой или нет.|
||[numberFormat](/javascript/api/excel/excel.chartaxisloadoptions#numberformat)|Представляет код формата для метки делений оси.|
||[корреспондирующей](/javascript/api/excel/excel.chartaxisloadoptions#offset)|Представляет расстояние между уровнями меток и расстояние между первым уровнем и линией оси. Значение должно быть целым числом от 0 до 1000.|
||[position](/javascript/api/excel/excel.chartaxisloadoptions#position)|Представляет указанное положение оси в месте, где ее пересекает другая ось. Дополнительные сведения см. в статье Excel. Чартаксиспоситион.|
||[Поситионат](/javascript/api/excel/excel.chartaxisloadoptions#positionat)|Представляет указанное положение оси в месте, где ее пересекает другая ось. Чтобы задать это свойство, следует использовать метод SetPositionAt(double).|
||[textOrientation](/javascript/api/excel/excel.chartaxisloadoptions#textorientation)|Представляет ориентацию текста для метки делений оси. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[Сетформула (формула: строка)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Строковое значение, представляющее формулу заголовка оси диаграммы с использованием нотации стиля A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[вокруг](/javascript/api/excel/excel.chartaxistitleformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Представляет форматирование заливки диаграммы.|
|[Чартаксиститлеформатдата](/javascript/api/excel/excel.chartaxistitleformatdata)|[вокруг](/javascript/api/excel/excel.chartaxistitleformatdata#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[Чартаксиститлеформатлоадоптионс](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[вокруг](/javascript/api/excel/excel.chartaxistitleformatloadoptions#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[Чартаксиститлеформатупдатедата](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[вокруг](/javascript/api/excel/excel.chartaxistitleformatupdatedata#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[Чартаксисупдатедата](/javascript/api/excel/excel.chartaxisupdatedata)|[ориентации](/javascript/api/excel/excel.chartaxisupdatedata#alignment)|Представляет выравнивание для указанной метки делений оси. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[Исбетвинкатегориес](/javascript/api/excel/excel.chartaxisupdatedata#isbetweencategories)|Указывает, пересекает ли ось значений ось категорий между категориями.|
||[Уровневые](/javascript/api/excel/excel.chartaxisupdatedata#multilevel)|Указывает, является ли ось многоуровневой или нет.|
||[numberFormat](/javascript/api/excel/excel.chartaxisupdatedata#numberformat)|Представляет код формата для метки делений оси.|
||[корреспондирующей](/javascript/api/excel/excel.chartaxisupdatedata#offset)|Представляет расстояние между уровнями меток и расстояние между первым уровнем и линией оси. Значение должно быть целым числом от 0 до 1000.|
||[position](/javascript/api/excel/excel.chartaxisupdatedata#position)|Представляет указанное положение оси в месте, где ее пересекает другая ось. Дополнительные сведения см. в статье Excel. Чартаксиспоситион.|
||[textOrientation](/javascript/api/excel/excel.chartaxisupdatedata#textorientation)|Представляет ориентацию текста для метки делений оси. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Очищает формат границы элемента диаграммы.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Возникает при активации диаграммы.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Возникает при добавлении новой диаграммы на лист.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Возникает при отключении диаграммы.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Возникает при удалении диаграммы.|
|[Чартколлектионлоадоптионс](/javascript/api/excel/excel.chartcollectionloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartcollectionloadoptions#categorylabellevel)|Для каждого элемента в коллекции: Возвращает или задает константу перечисления Чарткатегорилабеллевел, ссылающуюся на|
||[displayBlanksAs](/javascript/api/excel/excel.chartcollectionloadoptions#displayblanksas)|Для каждого элемента в коллекции: Возвращает или задает способ отображения пустых ячеек на диаграмме. Для чтения и записи.|
||[plotArea](/javascript/api/excel/excel.chartcollectionloadoptions#plotarea)|Для каждого элемента в коллекции: представляет plotArea для диаграммы.|
||[plotBy](/javascript/api/excel/excel.chartcollectionloadoptions#plotby)|Для каждого элемента в коллекции: Возвращает или задает способ использования столбцов или строк в качестве рядов данных на диаграмме. Для чтения и записи.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartcollectionloadoptions#plotvisibleonly)|Для каждого элемента в коллекции: true, если отображаются только видимые ячейки.False, если отображаются как видимые, так и скрытые ячейки. Для чтения и записи.|
||[seriesNameLevel](/javascript/api/excel/excel.chartcollectionloadoptions#seriesnamelevel)|Для каждого элемента в коллекции: Возвращает или задает константу перечисления Чартсериеснамелевел, ссылающуюся на|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartcollectionloadoptions#showdatalabelsovermaximum)|Для каждого элемента в коллекции: указывает, показывать ли метки данных, если значение больше максимального значения на оси значений.|
||[style](/javascript/api/excel/excel.chartcollectionloadoptions#style)|Для каждого элемента в коллекции: Возвращает или задает стиль диаграммы для диаграммы. Для чтения и записи.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[categoryLabelLevel](/javascript/api/excel/excel.chartdata#categorylabellevel)|Возвращает или задает константу перечисления Чарткатегорилабеллевел, ссылающуюся на|
||[displayBlanksAs](/javascript/api/excel/excel.chartdata#displayblanksas)|Возвращает или задает способ отображения пустых ячеек на диаграмме. Для чтения и записи.|
||[plotArea](/javascript/api/excel/excel.chartdata#plotarea)|Представляет plotArea для диаграммы.|
||[plotBy](/javascript/api/excel/excel.chartdata#plotby)|Возвращает или задает способ использования столбцов или строк в качестве рядов данных на диаграмме. Для чтения и записи.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartdata#plotvisibleonly)|True, если отображаются только видимые ячейки.False, если отображаются как видимые, так и скрытые ячейки. Для чтения и записи.|
||[seriesNameLevel](/javascript/api/excel/excel.chartdata#seriesnamelevel)|Возвращает или задает константу перечисления Чартсериеснамелевел, ссылающуюся на|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartdata#showdatalabelsovermaximum)|Представляет, нужно ли отображать метки данных, если значение больше максимального на оси значений.|
||[style](/javascript/api/excel/excel.chartdata#style)|Возвращает или задает стиль для диаграммы. Для чтения и записи.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[Элемента](/javascript/api/excel/excel.chartdatalabel#autotext)|Логическое значение, указывающее на то, генерирует ли метка данных автоматически соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах).  Значение NULL, если метка данных диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Строковое значение, представляющее код формата для метки данных.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Представляет формат метки данных диаграммы.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Возвращает высоту метки данных диаграммы (в пунктах). Только для чтения. Значение NULL, если метка данных диаграммы не отображается.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Возвращает ширину метки данных диаграммы (в пунктах). Только для чтения. Значение NULL, если метка данных диаграммы не отображается.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Представляет ориентацию текста для метки данных диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах). Значение NULL, если метка данных диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чартдаталабелдата](/javascript/api/excel/excel.chartdatalabeldata)|[Элемента](/javascript/api/excel/excel.chartdatalabeldata#autotext)|Логическое значение, указывающее на то, генерирует ли метка данных автоматически соответствующий текст на основе контекста.|
||[format](/javascript/api/excel/excel.chartdatalabeldata#format)|Представляет формат метки данных диаграммы.|
||[formula](/javascript/api/excel/excel.chartdatalabeldata#formula)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[height](/javascript/api/excel/excel.chartdatalabeldata#height)|Возвращает высоту метки данных диаграммы (в пунктах). Только для чтения. Значение NULL, если метка данных диаграммы не отображается.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabeldata#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.chartdatalabeldata#left)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах).  Значение NULL, если метка данных диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabeldata#numberformat)|Строковое значение, представляющее код формата для метки данных.|
||[text](/javascript/api/excel/excel.chartdatalabeldata#text)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabeldata#textorientation)|Представляет ориентацию текста для метки данных диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.chartdatalabeldata#top)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах). Значение NULL, если метка данных диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabeldata#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
||[width](/javascript/api/excel/excel.chartdatalabeldata#width)|Возвращает ширину метки данных диаграммы (в пунктах). Только для чтения. Значение NULL, если метка данных диаграммы не отображается.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[вокруг](/javascript/api/excel/excel.chartdatalabelformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину. Только для чтения.|
|[Чартдаталабелформатдата](/javascript/api/excel/excel.chartdatalabelformatdata)|[вокруг](/javascript/api/excel/excel.chartdatalabelformatdata#border)|Представляет формат границы, включающий цвет, тип линии и толщину. Только для чтения.|
|[Чартдаталабелформатлоадоптионс](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[вокруг](/javascript/api/excel/excel.chartdatalabelformatloadoptions#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[Чартдаталабелформатупдатедата](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[вокруг](/javascript/api/excel/excel.chartdatalabelformatupdatedata#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[Чартдаталабеллоадоптионс](/javascript/api/excel/excel.chartdatalabelloadoptions)|[Элемента](/javascript/api/excel/excel.chartdatalabelloadoptions#autotext)|Логическое значение, указывающее на то, генерирует ли метка данных автоматически соответствующий текст на основе контекста.|
||[format](/javascript/api/excel/excel.chartdatalabelloadoptions#format)|Представляет формат метки данных диаграммы.|
||[formula](/javascript/api/excel/excel.chartdatalabelloadoptions#formula)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[height](/javascript/api/excel/excel.chartdatalabelloadoptions#height)|Возвращает высоту метки данных диаграммы (в пунктах). Только для чтения. Значение NULL, если метка данных диаграммы не отображается.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.chartdatalabelloadoptions#left)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах).  Значение NULL, если метка данных диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#numberformat)|Строковое значение, представляющее код формата для метки данных.|
||[text](/javascript/api/excel/excel.chartdatalabelloadoptions#text)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelloadoptions#textorientation)|Представляет ориентацию текста для метки данных диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.chartdatalabelloadoptions#top)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах). Значение NULL, если метка данных диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
||[width](/javascript/api/excel/excel.chartdatalabelloadoptions#width)|Возвращает ширину метки данных диаграммы (в пунктах). Только для чтения. Значение NULL, если метка данных диаграммы не отображается.|
|[Чартдаталабелупдатедата](/javascript/api/excel/excel.chartdatalabelupdatedata)|[Элемента](/javascript/api/excel/excel.chartdatalabelupdatedata#autotext)|Логическое значение, указывающее на то, генерирует ли метка данных автоматически соответствующий текст на основе контекста.|
||[format](/javascript/api/excel/excel.chartdatalabelupdatedata#format)|Представляет формат метки данных диаграммы.|
||[formula](/javascript/api/excel/excel.chartdatalabelupdatedata#formula)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.chartdatalabelupdatedata#left)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах).  Значение NULL, если метка данных диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#numberformat)|Строковое значение, представляющее код формата для метки данных.|
||[text](/javascript/api/excel/excel.chartdatalabelupdatedata#text)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelupdatedata#textorientation)|Представляет ориентацию текста для метки данных диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.chartdatalabelupdatedata#top)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах). Значение NULL, если метка данных диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[Элемента](/javascript/api/excel/excel.chartdatalabels#autotext)|Указывает, генерируют ли метки данных автоматически соответствующий текст на основе контекста.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Представляет код формата для меток данных.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Представляет ориентацию текста для меток данных. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чартдаталабелсдата](/javascript/api/excel/excel.chartdatalabelsdata)|[Элемента](/javascript/api/excel/excel.chartdatalabelsdata#autotext)|Указывает, генерируют ли метки данных автоматически соответствующий текст на основе контекста.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsdata#numberformat)|Представляет код формата для меток данных.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsdata#textorientation)|Представляет ориентацию текста для меток данных. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чартдаталабелслоадоптионс](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[Элемента](/javascript/api/excel/excel.chartdatalabelsloadoptions#autotext)|Указывает, генерируют ли метки данных автоматически соответствующий текст на основе контекста.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#numberformat)|Представляет код формата для меток данных.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsloadoptions#textorientation)|Представляет ориентацию текста для меток данных. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чартдаталабелсупдатедата](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[Элемента](/javascript/api/excel/excel.chartdatalabelsupdatedata#autotext)|Указывает, генерируют ли метки данных автоматически соответствующий текст на основе контекста.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#numberformat)|Представляет код формата для меток данных.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsupdatedata#textorientation)|Представляет ориентацию текста для меток данных. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чартдеактиватедевентаргс](/javascript/api/excel/excel.chartdeactivatedeventargs)|[Чартид](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Получает идентификатор деактивированной диаграммы.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором деактивирована диаграмма.|
|[Чартделетедевентаргс](/javascript/api/excel/excel.chartdeletedeventargs)|[Чартид](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Получает идентификатор диаграммы, удаляемой с листа.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Получает идентификатор листа, в котором удаляется диаграмма.|
|[Чартлежендентри](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Представляет высоту объекта legendEntry в условных обозначениях диаграммы.|
||[индекс](/javascript/api/excel/excel.chartlegendentry#index)|Представляет индекс объекта legendEntry в условных обозначениях диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Представляет левую часть объекта legendEntry диаграммы.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Представляет верхнюю часть объекта legendEntry диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Представляет ширину объекта legendEntry в условных обозначениях диаграммы.|
|[Чартлежендентриколлектионлоадоптионс](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[height](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#height)|Для каждого элемента в коллекции: представляет высоту legendEntry в условных обозначениях диаграммы.|
||[индекс](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#index)|Для каждого элемента в коллекции: представляет индекс объекта legendEntry в условных обозначениях диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#left)|Для каждого элемента в коллекции: представляет собой левую часть диаграммы legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#top)|Для каждого элемента в коллекции — представляет верхнюю часть диаграммы legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#width)|Для каждого элемента в коллекции: представляет ширину legendEntry в условных обозначениях диаграммы.|
|[Чартлежендентридата](/javascript/api/excel/excel.chartlegendentrydata)|[height](/javascript/api/excel/excel.chartlegendentrydata#height)|Представляет высоту объекта legendEntry в условных обозначениях диаграммы.|
||[индекс](/javascript/api/excel/excel.chartlegendentrydata#index)|Представляет индекс объекта legendEntry в условных обозначениях диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentrydata#left)|Представляет левую часть объекта legendEntry диаграммы.|
||[top](/javascript/api/excel/excel.chartlegendentrydata#top)|Представляет верхнюю часть объекта legendEntry диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendentrydata#width)|Представляет ширину объекта legendEntry в условных обозначениях диаграммы.|
|[Чартлежендентрилоадоптионс](/javascript/api/excel/excel.chartlegendentryloadoptions)|[height](/javascript/api/excel/excel.chartlegendentryloadoptions#height)|Представляет высоту объекта legendEntry в условных обозначениях диаграммы.|
||[индекс](/javascript/api/excel/excel.chartlegendentryloadoptions#index)|Представляет индекс объекта legendEntry в условных обозначениях диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentryloadoptions#left)|Представляет левую часть объекта legendEntry диаграммы.|
||[top](/javascript/api/excel/excel.chartlegendentryloadoptions#top)|Представляет верхнюю часть объекта legendEntry диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendentryloadoptions#width)|Представляет ширину объекта legendEntry в условных обозначениях диаграммы.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[вокруг](/javascript/api/excel/excel.chartlegendformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину. Только для чтения.|
|[Чартлежендформатдата](/javascript/api/excel/excel.chartlegendformatdata)|[вокруг](/javascript/api/excel/excel.chartlegendformatdata#border)|Представляет формат границы, включающий цвет, тип линии и толщину. Только для чтения.|
|[Чартлежендформатлоадоптионс](/javascript/api/excel/excel.chartlegendformatloadoptions)|[вокруг](/javascript/api/excel/excel.chartlegendformatloadoptions#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[Чартлежендформатупдатедата](/javascript/api/excel/excel.chartlegendformatupdatedata)|[вокруг](/javascript/api/excel/excel.chartlegendformatupdatedata#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[Чартлоадоптионс](/javascript/api/excel/excel.chartloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartloadoptions#categorylabellevel)|Возвращает или задает константу перечисления Чарткатегорилабеллевел, ссылающуюся на|
||[displayBlanksAs](/javascript/api/excel/excel.chartloadoptions#displayblanksas)|Возвращает или задает способ отображения пустых ячеек на диаграмме. Для чтения и записи.|
||[plotArea](/javascript/api/excel/excel.chartloadoptions#plotarea)|Представляет plotArea для диаграммы.|
||[plotBy](/javascript/api/excel/excel.chartloadoptions#plotby)|Возвращает или задает способ использования столбцов или строк в качестве рядов данных на диаграмме. Для чтения и записи.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartloadoptions#plotvisibleonly)|True, если отображаются только видимые ячейки.False, если отображаются как видимые, так и скрытые ячейки. Для чтения и записи.|
||[seriesNameLevel](/javascript/api/excel/excel.chartloadoptions#seriesnamelevel)|Возвращает или задает константу перечисления Чартсериеснамелевел, ссылающуюся на|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartloadoptions#showdatalabelsovermaximum)|Представляет, нужно ли отображать метки данных, если значение больше максимального на оси значений.|
||[style](/javascript/api/excel/excel.chartloadoptions#style)|Возвращает или задает стиль для диаграммы. Для чтения и записи.|
|[Чартплотареа](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Представляет значение высоты plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Представляет значение insideHeight для plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Представляет значение insideLeft для plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Представляет значение insideTop для plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Представляет значение insideWidth для plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Представляет левое значение plotArea.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Представляет положение plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Представляет форматирование для plotArea диаграммы.|
||[Set (Properties: Excel. Чартплотареа)](/javascript/api/excel/excel.chartplotarea#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартплотареаупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartplotarea#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Представляет верхнее значение plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Представляет значение ширины plotArea.|
|[Чартплотареадата](/javascript/api/excel/excel.chartplotareadata)|[format](/javascript/api/excel/excel.chartplotareadata#format)|Представляет форматирование для plotArea диаграммы.|
||[height](/javascript/api/excel/excel.chartplotareadata#height)|Представляет значение высоты plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotareadata#insideheight)|Представляет значение insideHeight для plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotareadata#insideleft)|Представляет значение insideLeft для plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotareadata#insidetop)|Представляет значение insideTop для plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotareadata#insidewidth)|Представляет значение insideWidth для plotArea.|
||[left](/javascript/api/excel/excel.chartplotareadata#left)|Представляет левое значение plotArea.|
||[position](/javascript/api/excel/excel.chartplotareadata#position)|Представляет положение plotArea.|
||[top](/javascript/api/excel/excel.chartplotareadata#top)|Представляет верхнее значение plotArea.|
||[width](/javascript/api/excel/excel.chartplotareadata#width)|Представляет значение ширины plotArea.|
|[Чартплотареаформат](/javascript/api/excel/excel.chartplotareaformat)|[вокруг](/javascript/api/excel/excel.chartplotareaformat#border)|Представляет атрибуты границы для plotArea диаграммы.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[Set (Properties: Excel. Чартплотареаформат)](/javascript/api/excel/excel.chartplotareaformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартплотареаформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartplotareaformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартплотареаформатдата](/javascript/api/excel/excel.chartplotareaformatdata)|[вокруг](/javascript/api/excel/excel.chartplotareaformatdata#border)|Представляет атрибуты границы для plotArea диаграммы.|
|[Чартплотареаформатлоадоптионс](/javascript/api/excel/excel.chartplotareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartplotareaformatloadoptions#$all)||
||[вокруг](/javascript/api/excel/excel.chartplotareaformatloadoptions#border)|Представляет атрибуты границы для plotArea диаграммы.|
|[Чартплотареаформатупдатедата](/javascript/api/excel/excel.chartplotareaformatupdatedata)|[вокруг](/javascript/api/excel/excel.chartplotareaformatupdatedata#border)|Представляет атрибуты границы для plotArea диаграммы.|
|[Чартплотареалоадоптионс](/javascript/api/excel/excel.chartplotarealoadoptions)|[$all](/javascript/api/excel/excel.chartplotarealoadoptions#$all)||
||[format](/javascript/api/excel/excel.chartplotarealoadoptions#format)|Представляет форматирование для plotArea диаграммы.|
||[height](/javascript/api/excel/excel.chartplotarealoadoptions#height)|Представляет значение высоты plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarealoadoptions#insideheight)|Представляет значение insideHeight для plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarealoadoptions#insideleft)|Представляет значение insideLeft для plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarealoadoptions#insidetop)|Представляет значение insideTop для plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarealoadoptions#insidewidth)|Представляет значение insideWidth для plotArea.|
||[left](/javascript/api/excel/excel.chartplotarealoadoptions#left)|Представляет левое значение plotArea.|
||[position](/javascript/api/excel/excel.chartplotarealoadoptions#position)|Представляет положение plotArea.|
||[top](/javascript/api/excel/excel.chartplotarealoadoptions#top)|Представляет верхнее значение plotArea.|
||[width](/javascript/api/excel/excel.chartplotarealoadoptions#width)|Представляет значение ширины plotArea.|
|[Чартплотареаупдатедата](/javascript/api/excel/excel.chartplotareaupdatedata)|[format](/javascript/api/excel/excel.chartplotareaupdatedata#format)|Представляет форматирование для plotArea диаграммы.|
||[height](/javascript/api/excel/excel.chartplotareaupdatedata#height)|Представляет значение высоты plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotareaupdatedata#insideheight)|Представляет значение insideHeight для plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotareaupdatedata#insideleft)|Представляет значение insideLeft для plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotareaupdatedata#insidetop)|Представляет значение insideTop для plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotareaupdatedata#insidewidth)|Представляет значение insideWidth для plotArea.|
||[left](/javascript/api/excel/excel.chartplotareaupdatedata#left)|Представляет левое значение plotArea.|
||[position](/javascript/api/excel/excel.chartplotareaupdatedata#position)|Представляет положение plotArea.|
||[top](/javascript/api/excel/excel.chartplotareaupdatedata#top)|Представляет верхнее значение plotArea.|
||[width](/javascript/api/excel/excel.chartplotareaupdatedata#width)|Представляет значение ширины plotArea.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Возвращает или задает группу для указанного ряда. Чтение и запись|
||[развертывани](/javascript/api/excel/excel.chartseries#explosion)|Возвращает или задает значение развертывания для сектора круговой или кольцевой диаграммы. Возвращает нуль (0) при отсутствии развертывания (верхушка сектора — в центре круговой диаграммы). Для чтения и записи.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Возвращает или задает угол первого сектора круговой или кольцевой диаграммы, в градусах (по часовой стрелке из вертикального положения). Применяется только к круговым, объемным круговым и кольцевым диаграммам. Может находиться в диапазоне от 0 до 360. Чтение и запись|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|Значение true, если Microsoft Excel инвертирует шаблон в элементе, когда он соответствует отрицательному числу. Для чтения и записи.|
||[перекрывающееся](/javascript/api/excel/excel.chartseries#overlap)|Указывает на расположение строк и столбцов. Может принимать значение от – 100 до 100. Применяется только к двумерным диаграммам и гистограммам. Для чтения и записи.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Представляет коллекцию всех dataLabels в ряду.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Возвращает или задает размер вторичного раздела круга круговой диаграммы либо линии круговой диаграммы в процентах от размера основной круговой диаграммы. Может находиться в диапазоне от 5 до 200. Для чтения и записи.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Возвращает или задает способ разделения двух разделов круга круговой диаграммы либо линии круговой диаграммы. Для чтения и записи.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|Значение true, если Microsoft Excel назначает разные цвета или шаблоны каждому маркеру данных. Диаграмма должна содержать только один ряд. Для чтения и записи.|
|[Чартсериесколлектионлоадоптионс](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriescollectionloadoptions#axisgroup)|Для каждого элемента в коллекции: Возвращает или задает группу для указанного ряда. Чтение и запись|
||[dataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#datalabels)|Для каждого элемента в коллекции: представляет коллекцию всех подписей данных в ряду.|
||[развертывани](/javascript/api/excel/excel.chartseriescollectionloadoptions#explosion)|Для каждого элемента в коллекции: Возвращает или задает значение развертывания для круговой диаграммы или сектора кольцевой диаграммы. Возвращает нуль (0) при отсутствии развертывания (верхушка сектора — в центре круговой диаграммы). Для чтения и записи.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriescollectionloadoptions#firstsliceangle)|Для каждого элемента в коллекции: Возвращает или задает угол первого сектора круговой диаграммы или кольцевой диаграммы в градусах (по часовой стрелке от вертикального). Применяется только к круговым, объемным круговым и кольцевым диаграммам. Может находиться в диапазоне от 0 до 360. Чтение и запись|
||[invertIfNegative](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertifnegative)|Для каждого элемента в коллекции: имеет значение true, если Microsoft Excel инвертирует шаблон элемента, если он соответствует отрицательному числу. Для чтения и записи.|
||[перекрывающееся](/javascript/api/excel/excel.chartseriescollectionloadoptions#overlap)|Для каждого элемента в коллекции: указывает, как располагаются полосы и столбцы. Может принимать значение от – 100 до 100. Применяется только к двумерным диаграммам и гистограммам. Для чтения и записи.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#secondplotsize)|Для каждого элемента в коллекции: Возвращает или задает размер вторичного раздела круговой диаграммы или круговой диаграммы в процентах от размера основной круговой диаграммы. Может находиться в диапазоне от 5 до 200. Для чтения и записи.|
||[splitType](/javascript/api/excel/excel.chartseriescollectionloadoptions#splittype)|Для каждого элемента в коллекции: Возвращает или задает способ разделения двух разделов круговой диаграммы или круговой диаграммы. Для чтения и записи.|
||[varyByCategories](/javascript/api/excel/excel.chartseriescollectionloadoptions#varybycategories)|Для каждого элемента в коллекции: имеет значение true, если Microsoft Excel назначает разные цвета или узоры для каждого маркера данных. Диаграмма должна содержать только один ряд. Для чтения и записи.|
|[Чартсериесдата](/javascript/api/excel/excel.chartseriesdata)|[axisGroup](/javascript/api/excel/excel.chartseriesdata#axisgroup)|Возвращает или задает группу для указанного ряда. Чтение и запись|
||[dataLabels](/javascript/api/excel/excel.chartseriesdata#datalabels)|Представляет коллекцию всех dataLabels в ряду.|
||[развертывани](/javascript/api/excel/excel.chartseriesdata#explosion)|Возвращает или задает значение развертывания для сектора круговой или кольцевой диаграммы. Возвращает нуль (0) при отсутствии развертывания (верхушка сектора — в центре круговой диаграммы). Для чтения и записи.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesdata#firstsliceangle)|Возвращает или задает угол первого сектора круговой или кольцевой диаграммы, в градусах (по часовой стрелке из вертикального положения). Применяется только к круговым, объемным круговым и кольцевым диаграммам. Может находиться в диапазоне от 0 до 360. Чтение и запись|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesdata#invertifnegative)|Значение true, если Microsoft Excel инвертирует шаблон в элементе, когда он соответствует отрицательному числу. Для чтения и записи.|
||[перекрывающееся](/javascript/api/excel/excel.chartseriesdata#overlap)|Указывает на расположение строк и столбцов. Может принимать значение от – 100 до 100. Применяется только к двумерным диаграммам и гистограммам. Для чтения и записи.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesdata#secondplotsize)|Возвращает или задает размер вторичного раздела круга круговой диаграммы либо линии круговой диаграммы в процентах от размера основной круговой диаграммы. Может находиться в диапазоне от 5 до 200. Для чтения и записи.|
||[splitType](/javascript/api/excel/excel.chartseriesdata#splittype)|Возвращает или задает способ разделения двух разделов круга круговой диаграммы либо линии круговой диаграммы. Для чтения и записи.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesdata#varybycategories)|Значение true, если Microsoft Excel назначает разные цвета или шаблоны каждому маркеру данных. Диаграмма должна содержать только один ряд. Для чтения и записи.|
|[Чартсериеслоадоптионс](/javascript/api/excel/excel.chartseriesloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriesloadoptions#axisgroup)|Возвращает или задает группу для указанного ряда. Чтение и запись|
||[dataLabels](/javascript/api/excel/excel.chartseriesloadoptions#datalabels)|Представляет коллекцию всех dataLabels в ряду.|
||[развертывани](/javascript/api/excel/excel.chartseriesloadoptions#explosion)|Возвращает или задает значение развертывания для сектора круговой или кольцевой диаграммы. Возвращает нуль (0) при отсутствии развертывания (верхушка сектора — в центре круговой диаграммы). Для чтения и записи.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesloadoptions#firstsliceangle)|Возвращает или задает угол первого сектора круговой или кольцевой диаграммы, в градусах (по часовой стрелке из вертикального положения). Применяется только к круговым, объемным круговым и кольцевым диаграммам. Может находиться в диапазоне от 0 до 360. Чтение и запись|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesloadoptions#invertifnegative)|Значение true, если Microsoft Excel инвертирует шаблон в элементе, когда он соответствует отрицательному числу. Для чтения и записи.|
||[перекрывающееся](/javascript/api/excel/excel.chartseriesloadoptions#overlap)|Указывает на расположение строк и столбцов. Может принимать значение от – 100 до 100. Применяется только к двумерным диаграммам и гистограммам. Для чтения и записи.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesloadoptions#secondplotsize)|Возвращает или задает размер вторичного раздела круга круговой диаграммы либо линии круговой диаграммы в процентах от размера основной круговой диаграммы. Может находиться в диапазоне от 5 до 200. Для чтения и записи.|
||[splitType](/javascript/api/excel/excel.chartseriesloadoptions#splittype)|Возвращает или задает способ разделения двух разделов круга круговой диаграммы либо линии круговой диаграммы. Для чтения и записи.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesloadoptions#varybycategories)|Значение true, если Microsoft Excel назначает разные цвета или шаблоны каждому маркеру данных. Диаграмма должна содержать только один ряд. Для чтения и записи.|
|[Чартсериесупдатедата](/javascript/api/excel/excel.chartseriesupdatedata)|[axisGroup](/javascript/api/excel/excel.chartseriesupdatedata#axisgroup)|Возвращает или задает группу для указанного ряда. Чтение и запись|
||[dataLabels](/javascript/api/excel/excel.chartseriesupdatedata#datalabels)|Представляет коллекцию всех dataLabels в ряду.|
||[развертывани](/javascript/api/excel/excel.chartseriesupdatedata#explosion)|Возвращает или задает значение развертывания для сектора круговой или кольцевой диаграммы. Возвращает нуль (0) при отсутствии развертывания (верхушка сектора — в центре круговой диаграммы). Для чтения и записи.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesupdatedata#firstsliceangle)|Возвращает или задает угол первого сектора круговой или кольцевой диаграммы, в градусах (по часовой стрелке из вертикального положения). Применяется только к круговым, объемным круговым и кольцевым диаграммам. Может находиться в диапазоне от 0 до 360. Чтение и запись|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesupdatedata#invertifnegative)|Значение true, если Microsoft Excel инвертирует шаблон в элементе, когда он соответствует отрицательному числу. Для чтения и записи.|
||[перекрывающееся](/javascript/api/excel/excel.chartseriesupdatedata#overlap)|Указывает на расположение строк и столбцов. Может принимать значение от – 100 до 100. Применяется только к двумерным диаграммам и гистограммам. Для чтения и записи.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesupdatedata#secondplotsize)|Возвращает или задает размер вторичного раздела круга круговой диаграммы либо линии круговой диаграммы в процентах от размера основной круговой диаграммы. Может находиться в диапазоне от 5 до 200. Для чтения и записи.|
||[splitType](/javascript/api/excel/excel.chartseriesupdatedata#splittype)|Возвращает или задает способ разделения двух разделов круга круговой диаграммы либо линии круговой диаграммы. Для чтения и записи.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesupdatedata#varybycategories)|Значение true, если Microsoft Excel назначает разные цвета или шаблоны каждому маркеру данных. Диаграмма должна содержать только один ряд. Для чтения и записи.|
|[Чарттрендлине](/javascript/api/excel/excel.charttrendline)|[Бакквардпериод](/javascript/api/excel/excel.charttrendline#backwardperiod)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[Форвардпериод](/javascript/api/excel/excel.charttrendline#forwardperiod)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[Клей](/javascript/api/excel/excel.charttrendline#label)|Представляет метку линии тренда диаграммы.|
||[Шовекуатион](/javascript/api/excel/excel.charttrendline#showequation)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[Шоврскуаред](/javascript/api/excel/excel.charttrendline#showrsquared)|Значение true, если величина достоверности аппроксимации для линии тренда отображается на диаграмме.|
|[Чарттрендлинеколлектионлоадоптионс](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[Бакквардпериод](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#backwardperiod)|Для каждого элемента в коллекции: представляет число периодов, на которые линия тренда расширяется обратно.|
||[Форвардпериод](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#forwardperiod)|Для каждого элемента в коллекции: представляет число периодов, на которые линия тренда расширяется вперед.|
||[Клей](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#label)|Для каждого элемента в коллекции: представляет метку линии тренда диаграммы.|
||[Шовекуатион](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showequation)|Для каждого элемента в коллекции: true, если формула для линии тренда отображается на диаграмме.|
||[Шоврскуаред](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showrsquared)|Для каждого элемента в коллекции: true, если R-квадрат для линии тренда отображается на диаграмме.|
|[Чарттрендлинедата](/javascript/api/excel/excel.charttrendlinedata)|[Бакквардпериод](/javascript/api/excel/excel.charttrendlinedata#backwardperiod)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[Форвардпериод](/javascript/api/excel/excel.charttrendlinedata#forwardperiod)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[Клей](/javascript/api/excel/excel.charttrendlinedata#label)|Представляет метку линии тренда диаграммы.|
||[Шовекуатион](/javascript/api/excel/excel.charttrendlinedata#showequation)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[Шоврскуаред](/javascript/api/excel/excel.charttrendlinedata#showrsquared)|Значение true, если величина достоверности аппроксимации для линии тренда отображается на диаграмме.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[Элемента](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Логическое значение, указывающее на то, генерирует ли метка линии тренда автоматически соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Строковое значение, представляющее формулу подписи линии тренда диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Представляет горизонтальное выравнивание для подписи линии тренда диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Представляет расстояние от левого края подписи линии тренда диаграммы до левого края области диаграммы (в пунктах). Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Строковое значение, представляющее код формата для подписи линии тренда.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Представляет формат подписи линии тренда диаграммы.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Возвращает высоту подписи линии тренда диаграммы (в пунктах). Только для чтения. Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Возвращает ширину подписи линии тренда диаграммы (в пунктах). Только для чтения. Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[Set (Properties: Excel. Чарттрендлинелабел)](/javascript/api/excel/excel.charttrendlinelabel#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чарттрендлинелабелупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.charttrendlinelabel#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Строка, представляющая текст подписи линии тренда на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Представляет ориентацию текста для подписи линии тренда диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Представляет расстояние от верхнего края подписи линии тренда диаграммы до верха области диаграммы (в пунктах). Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Представляет вертикальное выравнивание для подписи линии тренда диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чарттрендлинелабелдата](/javascript/api/excel/excel.charttrendlinelabeldata)|[Элемента](/javascript/api/excel/excel.charttrendlinelabeldata#autotext)|Логическое значение, указывающее на то, генерирует ли метка линии тренда автоматически соответствующий текст на основе контекста.|
||[format](/javascript/api/excel/excel.charttrendlinelabeldata#format)|Представляет формат подписи линии тренда диаграммы.|
||[formula](/javascript/api/excel/excel.charttrendlinelabeldata#formula)|Строковое значение, представляющее формулу подписи линии тренда диаграммы с использованием нотации стиля A1.|
||[height](/javascript/api/excel/excel.charttrendlinelabeldata#height)|Возвращает высоту подписи линии тренда диаграммы (в пунктах). Только для чтения. Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#horizontalalignment)|Представляет горизонтальное выравнивание для подписи линии тренда диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.charttrendlinelabeldata#left)|Представляет расстояние от левого края подписи линии тренда диаграммы до левого края области диаграммы (в пунктах). Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#numberformat)|Строковое значение, представляющее код формата для подписи линии тренда.|
||[text](/javascript/api/excel/excel.charttrendlinelabeldata#text)|Строка, представляющая текст подписи линии тренда на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabeldata#textorientation)|Представляет ориентацию текста для подписи линии тренда диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.charttrendlinelabeldata#top)|Представляет расстояние от верхнего края подписи линии тренда диаграммы до верха области диаграммы (в пунктах). Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#verticalalignment)|Представляет вертикальное выравнивание для подписи линии тренда диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
||[width](/javascript/api/excel/excel.charttrendlinelabeldata#width)|Возвращает ширину подписи линии тренда диаграммы (в пунктах). Только для чтения. Значение NULL, если подпись линии тренда диаграммы не отображается.|
|[Чарттрендлинелабелформат](/javascript/api/excel/excel.charttrendlinelabelformat)|[вокруг](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Представляет формат заливки для текущей подписи линии тренда диаграммы.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Представляет атрибуты шрифта (имя, размер, цвет и т. д.) для подписи линии тренда диаграммы.|
||[Set (Properties: Excel. Чарттрендлинелабелформат)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чарттрендлинелабелформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чарттрендлинелабелформатдата](/javascript/api/excel/excel.charttrendlinelabelformatdata)|[вокруг](/javascript/api/excel/excel.charttrendlinelabelformatdata#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatdata#font)|Представляет атрибуты шрифта (имя, размер, цвет и т. д.) для подписи линии тренда диаграммы.|
|[Чарттрендлинелабелформатлоадоптионс](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#$all)||
||[вокруг](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#font)|Представляет атрибуты шрифта (имя, размер, цвет и т. д.) для подписи линии тренда диаграммы.|
|[Чарттрендлинелабелформатупдатедата](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata)|[вокруг](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#font)|Представляет атрибуты шрифта (имя, размер, цвет и т. д.) для подписи линии тренда диаграммы.|
|[Чарттрендлинелабеллоадоптионс](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelloadoptions#$all)||
||[Элемента](/javascript/api/excel/excel.charttrendlinelabelloadoptions#autotext)|Логическое значение, указывающее на то, генерирует ли метка линии тренда автоматически соответствующий текст на основе контекста.|
||[format](/javascript/api/excel/excel.charttrendlinelabelloadoptions#format)|Представляет формат подписи линии тренда диаграммы.|
||[formula](/javascript/api/excel/excel.charttrendlinelabelloadoptions#formula)|Строковое значение, представляющее формулу подписи линии тренда диаграммы с использованием нотации стиля A1.|
||[height](/javascript/api/excel/excel.charttrendlinelabelloadoptions#height)|Возвращает высоту подписи линии тренда диаграммы (в пунктах). Только для чтения. Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#horizontalalignment)|Представляет горизонтальное выравнивание для подписи линии тренда диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.charttrendlinelabelloadoptions#left)|Представляет расстояние от левого края подписи линии тренда диаграммы до левого края области диаграммы (в пунктах). Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#numberformat)|Строковое значение, представляющее код формата для подписи линии тренда.|
||[text](/javascript/api/excel/excel.charttrendlinelabelloadoptions#text)|Строка, представляющая текст подписи линии тренда на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelloadoptions#textorientation)|Представляет ориентацию текста для подписи линии тренда диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.charttrendlinelabelloadoptions#top)|Представляет расстояние от верхнего края подписи линии тренда диаграммы до верха области диаграммы (в пунктах). Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#verticalalignment)|Представляет вертикальное выравнивание для подписи линии тренда диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
||[width](/javascript/api/excel/excel.charttrendlinelabelloadoptions#width)|Возвращает ширину подписи линии тренда диаграммы (в пунктах). Только для чтения. Значение NULL, если подпись линии тренда диаграммы не отображается.|
|[Чарттрендлинелабелупдатедата](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[Элемента](/javascript/api/excel/excel.charttrendlinelabelupdatedata#autotext)|Логическое значение, указывающее на то, генерирует ли метка линии тренда автоматически соответствующий текст на основе контекста.|
||[format](/javascript/api/excel/excel.charttrendlinelabelupdatedata#format)|Представляет формат подписи линии тренда диаграммы.|
||[formula](/javascript/api/excel/excel.charttrendlinelabelupdatedata#formula)|Строковое значение, представляющее формулу подписи линии тренда диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#horizontalalignment)|Представляет горизонтальное выравнивание для подписи линии тренда диаграммы. Дополнительные сведения см. в статье Excel. Чарттекссоризонталалигнмент.|
||[left](/javascript/api/excel/excel.charttrendlinelabelupdatedata#left)|Представляет расстояние от левого края подписи линии тренда диаграммы до левого края области диаграммы (в пунктах). Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#numberformat)|Строковое значение, представляющее код формата для подписи линии тренда.|
||[text](/javascript/api/excel/excel.charttrendlinelabelupdatedata#text)|Строка, представляющая текст подписи линии тренда на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelupdatedata#textorientation)|Представляет ориентацию текста для подписи линии тренда диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.charttrendlinelabelupdatedata#top)|Представляет расстояние от верхнего края подписи линии тренда диаграммы до верха области диаграммы (в пунктах). Значение NULL, если подпись линии тренда диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#verticalalignment)|Представляет вертикальное выравнивание для подписи линии тренда диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чарттрендлинелоадоптионс](/javascript/api/excel/excel.charttrendlineloadoptions)|[Бакквардпериод](/javascript/api/excel/excel.charttrendlineloadoptions#backwardperiod)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[Форвардпериод](/javascript/api/excel/excel.charttrendlineloadoptions#forwardperiod)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[Клей](/javascript/api/excel/excel.charttrendlineloadoptions#label)|Представляет метку линии тренда диаграммы.|
||[Шовекуатион](/javascript/api/excel/excel.charttrendlineloadoptions#showequation)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[Шоврскуаред](/javascript/api/excel/excel.charttrendlineloadoptions#showrsquared)|Значение true, если величина достоверности аппроксимации для линии тренда отображается на диаграмме.|
|[Чарттрендлинеупдатедата](/javascript/api/excel/excel.charttrendlineupdatedata)|[Бакквардпериод](/javascript/api/excel/excel.charttrendlineupdatedata#backwardperiod)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[Форвардпериод](/javascript/api/excel/excel.charttrendlineupdatedata#forwardperiod)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[Клей](/javascript/api/excel/excel.charttrendlineupdatedata#label)|Представляет метку линии тренда диаграммы.|
||[Шовекуатион](/javascript/api/excel/excel.charttrendlineupdatedata#showequation)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[Шоврскуаред](/javascript/api/excel/excel.charttrendlineupdatedata#showrsquared)|Значение true, если величина достоверности аппроксимации для линии тренда отображается на диаграмме.|
|[Чартупдатедата](/javascript/api/excel/excel.chartupdatedata)|[categoryLabelLevel](/javascript/api/excel/excel.chartupdatedata#categorylabellevel)|Возвращает или задает константу перечисления Чарткатегорилабеллевел, ссылающуюся на|
||[displayBlanksAs](/javascript/api/excel/excel.chartupdatedata#displayblanksas)|Возвращает или задает способ отображения пустых ячеек на диаграмме. Для чтения и записи.|
||[plotArea](/javascript/api/excel/excel.chartupdatedata#plotarea)|Представляет plotArea для диаграммы.|
||[plotBy](/javascript/api/excel/excel.chartupdatedata#plotby)|Возвращает или задает способ использования столбцов или строк в качестве рядов данных на диаграмме. Для чтения и записи.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartupdatedata#plotvisibleonly)|True, если отображаются только видимые ячейки.False, если отображаются как видимые, так и скрытые ячейки. Для чтения и записи.|
||[seriesNameLevel](/javascript/api/excel/excel.chartupdatedata#seriesnamelevel)|Возвращает или задает константу перечисления Чартсериеснамелевел, ссылающуюся на|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartupdatedata#showdatalabelsovermaximum)|Представляет, нужно ли отображать метки данных, если значение больше максимального на оси значений.|
||[style](/javascript/api/excel/excel.chartupdatedata#style)|Возвращает или задает стиль для диаграммы. Для чтения и записи.|
|[Кустомдатавалидатион](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Формула проверки настраиваемых данных. При этом создаются специальные правила ввода, такие как предотвращение дублирования или ограничение суммы в диапазоне ячеек.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Числовой формат DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Положение DataPivotHierarchy.|
||[поле](/javascript/api/excel/excel.datapivothierarchy#field)|Возвращает сводные поля, связанные с DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|Идентификатор DataPivotHierarchy.|
||[Set (Properties: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchy#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Датапивосиерарчюпдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.datapivothierarchy#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[Сеттодефаулт ()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Сбрасывает DataPivotHierarchy до значений по умолчанию.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Определяет, должны ли данные отображаться как конкретные суммарные вычисления или нет.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Определяет, следует ли отображать все элементы DataPivotHierarchy.|
|[Датапивосиерарчиколлектион](/javascript/api/excel/excel.datapivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Получает DataPivotHierarchy по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Получает DataPivotHierarchy по имени. Если DataPivotHierarchy не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[Remove (DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[Датапивосиерарчиколлектионлоадоптионс](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#$all)||
||[поле](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#field)|Для каждого элемента в коллекции: возвращает PivotFields, связанный с DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#id)|Для каждого элемента в коллекции: ID объекта DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#name)|Для каждого элемента в коллекции: имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#numberformat)|Для каждого элемента в коллекции: числовой формат объекта DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#position)|Для каждого элемента в коллекции: положение DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#showas)|Для каждого элемента в коллекции: определяет, следует ли отображать данные в виде определенного сводного вычисления или нет.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#summarizeby)|Для каждого элемента в коллекции: определяет, нужно ли отображать все элементы объекта DataPivotHierarchy.|
|[Датапивосиерарчидата](/javascript/api/excel/excel.datapivothierarchydata)|[поле](/javascript/api/excel/excel.datapivothierarchydata#field)|Возвращает сводные поля, связанные с DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchydata#id)|Идентификатор DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchydata#name)|Имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchydata#numberformat)|Числовой формат DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchydata#position)|Положение DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchydata#showas)|Определяет, должны ли данные отображаться как конкретные суммарные вычисления или нет.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchydata#summarizeby)|Определяет, следует ли отображать все элементы DataPivotHierarchy.|
|[Датапивосиерарчилоадоптионс](/javascript/api/excel/excel.datapivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchyloadoptions#$all)||
||[поле](/javascript/api/excel/excel.datapivothierarchyloadoptions#field)|Возвращает сводные поля, связанные с DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchyloadoptions#id)|Идентификатор DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchyloadoptions#name)|Имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyloadoptions#numberformat)|Числовой формат DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchyloadoptions#position)|Положение DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchyloadoptions#showas)|Определяет, должны ли данные отображаться как конкретные суммарные вычисления или нет.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyloadoptions#summarizeby)|Определяет, следует ли отображать все элементы DataPivotHierarchy.|
|[Датапивосиерарчюпдатедата](/javascript/api/excel/excel.datapivothierarchyupdatedata)|[поле](/javascript/api/excel/excel.datapivothierarchyupdatedata#field)|Возвращает сводные поля, связанные с DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchyupdatedata#name)|Имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyupdatedata#numberformat)|Числовой формат DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchyupdatedata#position)|Положение DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchyupdatedata#showas)|Определяет, должны ли данные отображаться как конкретные суммарные вычисления или нет.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyupdatedata#summarizeby)|Определяет, следует ли отображать все элементы DataPivotHierarchy.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Очищает проверку данных из текущего диапазона.|
||[Ерроралерт](/javascript/api/excel/excel.datavalidation#erroralert)|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|
||[Игноребланкс](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Игнорировать пустые ячейки: проверка данных не будет выполняться в пустых ячейках, по умолчанию используется значение true.|
||[сообщение](/javascript/api/excel/excel.datavalidation#prompt)|Выдавать запрос при выборе пользователем ячейки.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Тип проверки данных, подробные сведения см. в статье Excel.DataValidationType.|
||[верно](/javascript/api/excel/excel.datavalidation#valid)|Указывает, являются ли все значения ячеек допустимыми в соответствии с правилами проверки данных.|
||[правила](/javascript/api/excel/excel.datavalidation#rule)|Правило проверки данных, которое содержит различные типы условий проверки данных.|
||[Set (Properties: Excel. IsValid)](/javascript/api/excel/excel.datavalidation#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Датавалидатионупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.datavalidation#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Датавалидатиондата](/javascript/api/excel/excel.datavalidationdata)|[Ерроралерт](/javascript/api/excel/excel.datavalidationdata#erroralert)|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|
||[Игноребланкс](/javascript/api/excel/excel.datavalidationdata#ignoreblanks)|Игнорировать пустые ячейки: проверка данных не будет выполняться в пустых ячейках, по умолчанию используется значение true.|
||[сообщение](/javascript/api/excel/excel.datavalidationdata#prompt)|Выдавать запрос при выборе пользователем ячейки.|
||[правила](/javascript/api/excel/excel.datavalidationdata#rule)|Правило проверки данных, которое содержит различные типы условий проверки данных.|
||[type](/javascript/api/excel/excel.datavalidationdata#type)|Тип проверки данных, подробные сведения см. в статье Excel.DataValidationType.|
||[верно](/javascript/api/excel/excel.datavalidationdata#valid)|Указывает, являются ли все значения ячеек допустимыми в соответствии с правилами проверки данных.|
|[Датавалидатионерроралерт](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Представляет предупреждающее сообщение об ошибке.|
||[Шовалерт](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Определяет, показывать ли диалоговое окно с предупреждением об ошибке или нет, если пользователь вводит неверные данные. Значение по умолчанию: true.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Представляет тип предупреждения проверки данных, подробные сведения см. в статье Excel.DataValidationAlertStyle.|
||[заголовок](/javascript/api/excel/excel.datavalidationerroralert#title)|Представляет заголовок диалогового окна предупреждения об ошибке.|
|[Датавалидатионлоадоптионс](/javascript/api/excel/excel.datavalidationloadoptions)|[$all](/javascript/api/excel/excel.datavalidationloadoptions#$all)||
||[Ерроралерт](/javascript/api/excel/excel.datavalidationloadoptions#erroralert)|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|
||[Игноребланкс](/javascript/api/excel/excel.datavalidationloadoptions#ignoreblanks)|Игнорировать пустые ячейки: проверка данных не будет выполняться в пустых ячейках, по умолчанию используется значение true.|
||[сообщение](/javascript/api/excel/excel.datavalidationloadoptions#prompt)|Выдавать запрос при выборе пользователем ячейки.|
||[правила](/javascript/api/excel/excel.datavalidationloadoptions#rule)|Правило проверки данных, которое содержит различные типы условий проверки данных.|
||[type](/javascript/api/excel/excel.datavalidationloadoptions#type)|Тип проверки данных, подробные сведения см. в статье Excel.DataValidationType.|
||[верно](/javascript/api/excel/excel.datavalidationloadoptions#valid)|Указывает, являются ли все значения ячеек допустимыми в соответствии с правилами проверки данных.|
|[Датавалидатионпромпт](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Представляет сообщение подсказки.|
||[Шовпромпт](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Определяет, показывать ли подсказку, когда пользователь выбирает ячейку с проверкой данных.|
||[заголовок](/javascript/api/excel/excel.datavalidationprompt#title)|Представляет заголовок подсказки.|
|[Датавалидатионруле](/javascript/api/excel/excel.datavalidationrule)|[собственный](/javascript/api/excel/excel.datavalidationrule#custom)|Условия проверки настраиваемых данных.|
||[дата](/javascript/api/excel/excel.datavalidationrule#date)|Условия проверки данных даты.|
||[числе](/javascript/api/excel/excel.datavalidationrule#decimal)|Условия проверки десятичных данных.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Условия проверки данных списка.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Условия проверки данных TextLength.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Условия проверки данных времени.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Условия проверки данных WholeNumber.|
|[Датавалидатионупдатедата](/javascript/api/excel/excel.datavalidationupdatedata)|[Ерроралерт](/javascript/api/excel/excel.datavalidationupdatedata#erroralert)|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|
||[Игноребланкс](/javascript/api/excel/excel.datavalidationupdatedata#ignoreblanks)|Игнорировать пустые ячейки: проверка данных не будет выполняться в пустых ячейках, по умолчанию используется значение true.|
||[сообщение](/javascript/api/excel/excel.datavalidationupdatedata#prompt)|Выдавать запрос при выборе пользователем ячейки.|
||[правила](/javascript/api/excel/excel.datavalidationupdatedata#rule)|Правило проверки данных, которое содержит различные типы условий проверки данных.|
|[Датетимедатавалидатион](/javascript/api/excel/excel.datetimedatavalidation)|[Formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Задает правый операнд, если для свойства operator задан бинарный оператор, такой как GreaterThan (левый операнд — это значение, которое пользователь пытается ввести в ячейку). С помощью операторов тернарного между и Нотбетвин задает нижнюю границу операнда.|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|С помощью операторов тернарного между и Нотбетвин указывает верхнюю границу операнда. Не используется с двоичными операторами, например GreaterThan.|
||[or](/javascript/api/excel/excel.datetimedatavalidation#operator)|Оператор, используемый для проверки данных.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[Енаблемултиплефилтеритемс](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Определяет, следует ли разрешить несколько элементов фильтра.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Положение FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Возвращает сводные поля, связанные с FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|Идентификатор FilterPivotHierarchy.|
||[Set (Properties: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchy#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Филтерпивосиерарчюпдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.filterpivothierarchy#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[Сеттодефаулт ()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Сбрасывает FilterPivotHierarchy до значений по умолчанию.|
|[Филтерпивосиерарчиколлектион](/javascript/api/excel/excel.filterpivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси. Если иерархия присутствует в другом месте строки, столбца,|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Получает FilterPivotHierarchy по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Получает FilterPivotHierarchy по имени. Если FilterPivotHierarchy не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[Remove (filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[Филтерпивосиерарчиколлектионлоадоптионс](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#$all)||
||[Енаблемултиплефилтеритемс](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#enablemultiplefilteritems)|Для каждого элемента в коллекции: определяет, разрешено ли использовать несколько элементов фильтра.|
||[id](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#id)|Для каждого элемента в коллекции: ID объекта FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#name)|Для каждого элемента в коллекции: имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#position)|Для каждого элемента в коллекции: положение FilterPivotHierarchy.|
|[Филтерпивосиерарчидата](/javascript/api/excel/excel.filterpivothierarchydata)|[Енаблемултиплефилтеритемс](/javascript/api/excel/excel.filterpivothierarchydata#enablemultiplefilteritems)|Определяет, следует ли разрешить несколько элементов фильтра.|
||[fields](/javascript/api/excel/excel.filterpivothierarchydata#fields)|Возвращает сводные поля, связанные с FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchydata#id)|Идентификатор FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchydata#name)|Имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchydata#position)|Положение FilterPivotHierarchy.|
|[Филтерпивосиерарчилоадоптионс](/javascript/api/excel/excel.filterpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchyloadoptions#$all)||
||[Енаблемултиплефилтеритемс](/javascript/api/excel/excel.filterpivothierarchyloadoptions#enablemultiplefilteritems)|Определяет, следует ли разрешить несколько элементов фильтра.|
||[id](/javascript/api/excel/excel.filterpivothierarchyloadoptions#id)|Идентификатор FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchyloadoptions#name)|Имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchyloadoptions#position)|Положение FilterPivotHierarchy.|
|[Филтерпивосиерарчюпдатедата](/javascript/api/excel/excel.filterpivothierarchyupdatedata)|[Енаблемултиплефилтеритемс](/javascript/api/excel/excel.filterpivothierarchyupdatedata#enablemultiplefilteritems)|Определяет, следует ли разрешить несколько элементов фильтра.|
||[name](/javascript/api/excel/excel.filterpivothierarchyupdatedata#name)|Имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchyupdatedata#position)|Положение FilterPivotHierarchy.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Отображает или не отображает список в раскрывающемся меню ячейки, по умолчанию используется значение true.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Источник списка для проверки данных|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Имя сводного поля.|
||[id](/javascript/api/excel/excel.pivotfield#id)|Идентификатор сводного поля.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Возвращает PivotItems, состоящий из PivotField.|
||[Set (Properties: Excel. PivotField)](/javascript/api/excel/excel.pivotfield#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Пивотфиелдупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.pivotfield#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Определяет, следует ли отображать все элементы сводного поля.|
||[Сортбилабелс (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Сортирует сводное поле. Если указан параметр DataPivotHierarchy, то сортировка будет применяться на его основе, в ином случае сортировка будет основана на самом сводном поле.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Промежуточные итоги сводного поля.|
|[Пивотфиелдколлектион](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Получает количество полей Pivot в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Получает объект PivotField по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Получает PivotField по имени. Если PivotField не существует, вернет пустой объект.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Пивотфиелдколлектионлоадоптионс](/javascript/api/excel/excel.pivotfieldcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#id)|Для каждого элемента в коллекции: ID объекта PivotField.|
||[name](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#name)|Для каждого элемента в коллекции: имя PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#showallitems)|Для каждого элемента в коллекции: определяет, нужно ли отображать все элементы объекта PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#subtotals)|Для каждого элемента в коллекции: подытоги объекта PivotField.|
|[Пивотфиелддата](/javascript/api/excel/excel.pivotfielddata)|[id](/javascript/api/excel/excel.pivotfielddata#id)|Идентификатор сводного поля.|
||[items](/javascript/api/excel/excel.pivotfielddata#items)|Возвращает сводные поля, связанные со сводным полем.|
||[name](/javascript/api/excel/excel.pivotfielddata#name)|Имя сводного поля.|
||[showAllItems](/javascript/api/excel/excel.pivotfielddata#showallitems)|Определяет, следует ли отображать все элементы сводного поля.|
||[subtotals](/javascript/api/excel/excel.pivotfielddata#subtotals)|Промежуточные итоги сводного поля.|
|[Пивотфиелдлоадоптионс](/javascript/api/excel/excel.pivotfieldloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldloadoptions#id)|Идентификатор сводного поля.|
||[name](/javascript/api/excel/excel.pivotfieldloadoptions#name)|Имя сводного поля.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldloadoptions#showallitems)|Определяет, следует ли отображать все элементы сводного поля.|
||[subtotals](/javascript/api/excel/excel.pivotfieldloadoptions#subtotals)|Промежуточные итоги сводного поля.|
|[Пивотфиелдупдатедата](/javascript/api/excel/excel.pivotfieldupdatedata)|[name](/javascript/api/excel/excel.pivotfieldupdatedata#name)|Имя сводного поля.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldupdatedata#showallitems)|Определяет, следует ли отображать все элементы сводного поля.|
||[subtotals](/javascript/api/excel/excel.pivotfieldupdatedata#subtotals)|Промежуточные итоги сводного поля.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Имя PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Возвращает сводные поля, связанные с PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|Идентификатор PivotHierarchy.|
||[Set (Properties: Excel. PivotHierarchy)](/javascript/api/excel/excel.pivothierarchy#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Пивосиерарчюпдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.pivothierarchy#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Пивосиерарчиколлектион](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Получает PivotHierarchy по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Получает PivotHierarchy по имени. Если PivotHierarchy не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Пивосиерарчиколлектионлоадоптионс](/javascript/api/excel/excel.pivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#id)|Для каждого элемента в коллекции: ID объекта PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#name)|Для каждого элемента в коллекции: имя PivotHierarchy.|
|[Пивосиерарчидата](/javascript/api/excel/excel.pivothierarchydata)|[fields](/javascript/api/excel/excel.pivothierarchydata#fields)|Возвращает сводные поля, связанные с PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchydata#id)|Идентификатор PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchydata#name)|Имя PivotHierarchy.|
|[Пивосиерарчилоадоптионс](/javascript/api/excel/excel.pivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchyloadoptions#id)|Идентификатор PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchyloadoptions#name)|Имя PivotHierarchy.|
|[Пивосиерарчюпдатедата](/javascript/api/excel/excel.pivothierarchyupdatedata)|[name](/javascript/api/excel/excel.pivothierarchyupdatedata#name)|Имя PivotHierarchy.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Определяет, развернут ли элемент для отображения дочерних элементов или же свернут, а дочерние элементы являются скрытыми.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Имя элемента сводной таблицы.|
||[id](/javascript/api/excel/excel.pivotitem#id)|Идентификатор элемента сводной таблицы.|
||[Set (Properties: Excel. PivotItem)](/javascript/api/excel/excel.pivotitem#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Пивотитемупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.pivotitem#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Определяет, отображается ли элемент сводной таблицы или нет.|
|[Пивотитемколлектион](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Получает количество элементов Pivot в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Получает объект PivotItem по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Получает PivotItem по имени. Если PivotItem не существует, вернет пустой объект.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Пивотитемколлектионлоадоптионс](/javascript/api/excel/excel.pivotitemcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotitemcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemcollectionloadoptions#id)|Для каждого элемента в коллекции: ID объекта PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitemcollectionloadoptions#isexpanded)|Для каждого элемента в коллекции: определяет, развернут ли элемент для отображения дочерних элементов или свернутый, и дочерние элементы скрыты.|
||[name](/javascript/api/excel/excel.pivotitemcollectionloadoptions#name)|Для каждого элемента в коллекции: имя PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemcollectionloadoptions#visible)|Для каждого элемента в коллекции: определяет, является ли PivotItem видимым.|
|[Пивотитемдата](/javascript/api/excel/excel.pivotitemdata)|[id](/javascript/api/excel/excel.pivotitemdata#id)|Идентификатор элемента сводной таблицы.|
||[isExpanded](/javascript/api/excel/excel.pivotitemdata#isexpanded)|Определяет, развернут ли элемент для отображения дочерних элементов или же свернут, а дочерние элементы являются скрытыми.|
||[name](/javascript/api/excel/excel.pivotitemdata#name)|Имя элемента сводной таблицы.|
||[visible](/javascript/api/excel/excel.pivotitemdata#visible)|Определяет, отображается ли элемент сводной таблицы или нет.|
|[Пивотитемлоадоптионс](/javascript/api/excel/excel.pivotitemloadoptions)|[$all](/javascript/api/excel/excel.pivotitemloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemloadoptions#id)|Идентификатор элемента сводной таблицы.|
||[isExpanded](/javascript/api/excel/excel.pivotitemloadoptions#isexpanded)|Определяет, развернут ли элемент для отображения дочерних элементов или же свернут, а дочерние элементы являются скрытыми.|
||[name](/javascript/api/excel/excel.pivotitemloadoptions#name)|Имя элемента сводной таблицы.|
||[visible](/javascript/api/excel/excel.pivotitemloadoptions#visible)|Определяет, отображается ли элемент сводной таблицы или нет.|
|[Пивотитемупдатедата](/javascript/api/excel/excel.pivotitemupdatedata)|[isExpanded](/javascript/api/excel/excel.pivotitemupdatedata#isexpanded)|Определяет, развернут ли элемент для отображения дочерних элементов или же свернут, а дочерние элементы являются скрытыми.|
||[name](/javascript/api/excel/excel.pivotitemupdatedata#name)|Имя элемента сводной таблицы.|
||[visible](/javascript/api/excel/excel.pivotitemupdatedata#visible)|Определяет, отображается ли элемент сводной таблицы или нет.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[Жетколумнлабелранже ()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Возвращает диапазон, где находятся названия столбцов сводной таблицы.|
||[Жетдатабодиранже ()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Возвращает диапазон, где находятся значения данных сводной таблицы.|
||[Жетфилтераксисранже ()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Возвращает диапазон области фильтра сводной таблицы.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Возвращает диапазон, в котором существует сводная таблица, за исключением области фильтра.|
||[Жетровлабелранже ()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Возвращает диапазон, где находятся названия строк сводной таблицы.|
||[Лайауттипе](/javascript/api/excel/excel.pivotlayout#layouttype)|Это свойство указывает PivotLayoutType всех полей в сводной таблице. Если поля имеют различные состояния, оно будет иметь значение NULL.|
||[Set (Properties: Excel. PivotLayout)](/javascript/api/excel/excel.pivotlayout#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Пивотлайаутупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.pivotlayout#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[Шовколумнграндтоталс](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для столбцов.|
||[Шовровграндтоталс](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для строк.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Это свойство указывает SubtotalLocationType всех полей в сводной таблице. Если поля имеют различные состояния, оно будет иметь значение NULL.|
|[Пивотлайаутдата](/javascript/api/excel/excel.pivotlayoutdata)|[Лайауттипе](/javascript/api/excel/excel.pivotlayoutdata#layouttype)|Это свойство указывает PivotLayoutType всех полей в сводной таблице. Если поля имеют различные состояния, оно будет иметь значение NULL.|
||[Шовколумнграндтоталс](/javascript/api/excel/excel.pivotlayoutdata#showcolumngrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для столбцов.|
||[Шовровграндтоталс](/javascript/api/excel/excel.pivotlayoutdata#showrowgrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для строк.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutdata#subtotallocation)|Это свойство указывает SubtotalLocationType всех полей в сводной таблице. Если поля имеют различные состояния, оно будет иметь значение NULL.|
|[Пивотлайаутлоадоптионс](/javascript/api/excel/excel.pivotlayoutloadoptions)|[$all](/javascript/api/excel/excel.pivotlayoutloadoptions#$all)||
||[Лайауттипе](/javascript/api/excel/excel.pivotlayoutloadoptions#layouttype)|Это свойство указывает PivotLayoutType всех полей в сводной таблице. Если поля имеют различные состояния, оно будет иметь значение NULL.|
||[Шовколумнграндтоталс](/javascript/api/excel/excel.pivotlayoutloadoptions#showcolumngrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для столбцов.|
||[Шовровграндтоталс](/javascript/api/excel/excel.pivotlayoutloadoptions#showrowgrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для строк.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutloadoptions#subtotallocation)|Это свойство указывает SubtotalLocationType всех полей в сводной таблице. Если поля имеют различные состояния, оно будет иметь значение NULL.|
|[Пивотлайаутупдатедата](/javascript/api/excel/excel.pivotlayoutupdatedata)|[Лайауттипе](/javascript/api/excel/excel.pivotlayoutupdatedata#layouttype)|Это свойство указывает PivotLayoutType всех полей в сводной таблице. Если поля имеют различные состояния, оно будет иметь значение NULL.|
||[Шовколумнграндтоталс](/javascript/api/excel/excel.pivotlayoutupdatedata#showcolumngrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для столбцов.|
||[Шовровграндтоталс](/javascript/api/excel/excel.pivotlayoutupdatedata#showrowgrandtotals)|Указывает, отображаются ли в отчете сводной таблицы общие итоги для строк.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutupdatedata#subtotallocation)|Это свойство указывает SubtotalLocationType всех полей в сводной таблице. Если поля имеют различные состояния, оно будет иметь значение NULL.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Удаляет сводную таблицу.|
||[Колумнхиерарчиес](/javascript/api/excel/excel.pivottable#columnhierarchies)|Иерархии сводных столбцов сводной таблицы.|
||[Иерархии](/javascript/api/excel/excel.pivottable#datahierarchies)|Иерархии сводных данных сводной таблицы.|
||[Филтерхиерарчиес](/javascript/api/excel/excel.pivottable#filterhierarchies)|Иерархии сводных фильтров сводной таблицы.|
||[иерархии](/javascript/api/excel/excel.pivottable#hierarchies)|Иерархии сводного документа сводной таблицы.|
||[макет](/javascript/api/excel/excel.pivottable#layout)|PivotLayout, описывающий макет и визуальную структуру сводной таблицы.|
||[Ровхиерарчиес](/javascript/api/excel/excel.pivottable#rowhierarchies)|Иерархии сводных строк сводной таблицы.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[Add (имя: строка, источник: таблица \| строк \| диапазона, назначение: строка \| диапазона)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Добавление сводной таблицы на основе указанных исходных данных и вставка ее в верхнюю левую ячейку конечного диапазона.|
|[Пивоттаблеколлектионлоадоптионс](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[макет](/javascript/api/excel/excel.pivottablecollectionloadoptions#layout)|Для каждого элемента в коллекции: PivotLayout, описывающий макет и визуальную структуру сводной таблицы.|
|[Пивоттабледата](/javascript/api/excel/excel.pivottabledata)|[Колумнхиерарчиес](/javascript/api/excel/excel.pivottabledata#columnhierarchies)|Иерархии сводных столбцов сводной таблицы.|
||[Иерархии](/javascript/api/excel/excel.pivottabledata#datahierarchies)|Иерархии сводных данных сводной таблицы.|
||[Филтерхиерарчиес](/javascript/api/excel/excel.pivottabledata#filterhierarchies)|Иерархии сводных фильтров сводной таблицы.|
||[иерархии](/javascript/api/excel/excel.pivottabledata#hierarchies)|Иерархии сводного документа сводной таблицы.|
||[Ровхиерарчиес](/javascript/api/excel/excel.pivottabledata#rowhierarchies)|Иерархии сводных строк сводной таблицы.|
|[Пивоттаблелоадоптионс](/javascript/api/excel/excel.pivottableloadoptions)|[макет](/javascript/api/excel/excel.pivottableloadoptions#layout)|PivotLayout, описывающий макет и визуальную структуру сводной таблицы.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Возвращает объект проверки данных.|
|[Ранжедата](/javascript/api/excel/excel.rangedata)|[dataValidation](/javascript/api/excel/excel.rangedata#datavalidation)|Возвращает объект проверки данных.|
|[Ранжелоадоптионс](/javascript/api/excel/excel.rangeloadoptions)|[dataValidation](/javascript/api/excel/excel.rangeloadoptions#datavalidation)|Возвращает объект проверки данных.|
|[Ранжеупдатедата](/javascript/api/excel/excel.rangeupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeupdatedata#datavalidation)|Возвращает объект проверки данных.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Положение RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Возвращает сводные поля, связанные с RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|Идентификатор RowColumnPivotHierarchy.|
||[Set (Properties: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Ровколумнпивосиерарчюпдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[Сеттодефаулт ()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Сбрасывает RowColumnPivotHierarchy до значений по умолчанию.|
|[Ровколумнпивосиерарчиколлектион](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси. Если иерархия присутствует в другом месте строки, столбца,|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Получает RowColumnPivotHierarchy по имени или идентификатору.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Получает RowColumnPivotHierarchy по имени. Если RowColumnPivotHierarchy не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[Remove (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[Ровколумнпивосиерарчиколлектионлоадоптионс](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#id)|Для каждого элемента в коллекции: ID объекта RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#name)|Для каждого элемента в коллекции: имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#position)|Для каждого элемента в коллекции: положение RowColumnPivotHierarchy.|
|[Ровколумнпивосиерарчидата](/javascript/api/excel/excel.rowcolumnpivothierarchydata)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchydata#fields)|Возвращает сводные поля, связанные с RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchydata#id)|Идентификатор RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchydata#name)|Имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchydata#position)|Положение RowColumnPivotHierarchy.|
|[Ровколумнпивосиерарчилоадоптионс](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#id)|Идентификатор RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#name)|Имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#position)|Положение RowColumnPivotHierarchy.|
|[Ровколумнпивосиерарчюпдатедата](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#name)|Имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#position)|Положение RowColumnPivotHierarchy.|
|[Полняющего](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Включение событий JavaScript в текущей области задач или контентной надстройке.|
|[Рунтимедата](/javascript/api/excel/excel.runtimedata)|[enableEvents](/javascript/api/excel/excel.runtimedata#enableevents)|Включение событий JavaScript в текущей области задач или контентной надстройке.|
|[Рунтимелоадоптионс](/javascript/api/excel/excel.runtimeloadoptions)|[enableEvents](/javascript/api/excel/excel.runtimeloadoptions#enableevents)|Включение событий JavaScript в текущей области задач или контентной надстройке.|
|[Рунтимеупдатедата](/javascript/api/excel/excel.runtimeupdatedata)|[enableEvents](/javascript/api/excel/excel.runtimeupdatedata#enableevents)|Включение событий JavaScript в текущей области задач или контентной надстройке.|
|[Шовасруле](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|Базовое сводное поле для обоснования расчета ShowAs, если применимо на основе типа ShowAsCalculation, в противном случае значение будет пустым.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|Базовый элемент для обоснования расчета ShowAs, если применимо на основе типа ShowAsCalculation, в противном случае значение будет пустым.|
||[пересчет](/javascript/api/excel/excel.showasrule#calculation)|Расчет ShowAs для использования в сводном поле данных. Дополнительные сведения см. в статье Excel. ShowAsCalculation.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста в ячейке установлено на равномерное распределение.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|Ориентация текста для стиля.|
|[Стилеколлектионлоадоптионс](/javascript/api/excel/excel.stylecollectionloadoptions)|[autoIndent](/javascript/api/excel/excel.stylecollectionloadoptions#autoindent)|Для каждого элемента в коллекции: указывает, отображается ли отступ текста автоматически, если для выравнивания текста в ячейке задано равное равномерное распределение.|
||[textOrientation](/javascript/api/excel/excel.stylecollectionloadoptions#textorientation)|Для каждого элемента в коллекции: ориентация текста для стиля.|
|[Стиледата](/javascript/api/excel/excel.styledata)|[autoIndent](/javascript/api/excel/excel.styledata#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста в ячейке установлено на равномерное распределение.|
||[textOrientation](/javascript/api/excel/excel.styledata#textorientation)|Ориентация текста для стиля.|
|[Стилелоадоптионс](/javascript/api/excel/excel.styleloadoptions)|[autoIndent](/javascript/api/excel/excel.styleloadoptions#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста в ячейке установлено на равномерное распределение.|
||[textOrientation](/javascript/api/excel/excel.styleloadoptions#textorientation)|Ориентация текста для стиля.|
|[Стилеупдатедата](/javascript/api/excel/excel.styleupdatedata)|[autoIndent](/javascript/api/excel/excel.styleupdatedata#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста в ячейке установлено на равномерное распределение.|
||[textOrientation](/javascript/api/excel/excel.styleupdatedata#textorientation)|Ориентация текста для стиля.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Если для свойства Automatic установлено значение true, все остальные значения будут игнорироваться при настройке промежуточных итогов.|
||[вычисления](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[Каунтнумберс](/javascript/api/excel/excel.subtotals#countnumbers)||
||[Max](/javascript/api/excel/excel.subtotals#max)||
||[минут](/javascript/api/excel/excel.subtotals#min)||
||[Продукция](/javascript/api/excel/excel.subtotals#product)||
||[Стандарддевиатион](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[Стандарддевиатионп](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[произведен](/javascript/api/excel/excel.subtotals#sum)||
||[различ](/javascript/api/excel/excel.subtotals#variance)||
||[Варианцеп](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|Возвращает числовой идентификатор.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область таблицы на конкретном листе.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область таблицы на конкретном листе. Может возвращать пустой объект.|
|[Таблеколлектионлоадоптионс](/javascript/api/excel/excel.tablecollectionloadoptions)|[legacyId](/javascript/api/excel/excel.tablecollectionloadoptions#legacyid)|Для каждого элемента в коллекции: Возвращает числовой идентификатор.|
|[TableData](/javascript/api/excel/excel.tabledata)|[legacyId](/javascript/api/excel/excel.tabledata#legacyid)|Возвращает числовой идентификатор.|
|[Таблелоадоптионс](/javascript/api/excel/excel.tableloadoptions)|[legacyId](/javascript/api/excel/excel.tableloadoptions#legacyid)|Возвращает числовой идентификатор.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|Значение true, если книга открыта в режиме только для чтения. Только для чтения.|
|[Воркбуккреатед](/javascript/api/excel/excel.workbookcreated)||[Воркбукдата](/javascript/api/excel/excel.workbookdata)|[readOnly](/javascript/api/excel/excel.workbookdata#readonly)|Значение true, если книга открыта в режиме только для чтения. Только для чтения.|
|[Воркбуклоадоптионс](/javascript/api/excel/excel.workbookloadoptions)|[readOnly](/javascript/api/excel/excel.workbookloadoptions#readonly)|Значение true, если книга открыта в режиме только для чтения. Только для чтения.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[oncalculated](/javascript/api/excel/excel.worksheet#oncalculated)|Возникает при вычислении листа.|
||[Шовгридлинес](/javascript/api/excel/excel.worksheet#showgridlines)|Получает или задает флаг линий сетки листа.|
||[Шовхеадингс](/javascript/api/excel/excel.worksheet#showheadings)|Получает или задает флаг заголовков листа.|
|[Воркшиткалкулатедевентаргс](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Получает идентификатор листа, который рассчитывается.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область конкретного листа. Может возвращать пустой объект.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[oncalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Возникает при вычислении любого листа в книге.|
|[Воркшитколлектионлоадоптионс](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[Шовгридлинес](/javascript/api/excel/excel.worksheetcollectionloadoptions#showgridlines)|Для каждого элемента в коллекции: Получает или задает флаг сетки листа.|
||[Шовхеадингс](/javascript/api/excel/excel.worksheetcollectionloadoptions#showheadings)|Для каждого элемента в коллекции: Получает или задает флаг заголовков листа.|
|[Воркшитдата](/javascript/api/excel/excel.worksheetdata)|[Шовгридлинес](/javascript/api/excel/excel.worksheetdata#showgridlines)|Получает или задает флаг линий сетки листа.|
||[Шовхеадингс](/javascript/api/excel/excel.worksheetdata#showheadings)|Получает или задает флаг заголовков листа.|
|[Воркшитлоадоптионс](/javascript/api/excel/excel.worksheetloadoptions)|[Шовгридлинес](/javascript/api/excel/excel.worksheetloadoptions#showgridlines)|Получает или задает флаг линий сетки листа.|
||[Шовхеадингс](/javascript/api/excel/excel.worksheetloadoptions#showheadings)|Получает или задает флаг заголовков листа.|
|[Воркшитупдатедата](/javascript/api/excel/excel.worksheetupdatedata)|[Шовгридлинес](/javascript/api/excel/excel.worksheetupdatedata#showgridlines)|Получает или задает флаг линий сетки листа.|
||[Шовхеадингс](/javascript/api/excel/excel.worksheetupdatedata#showheadings)|Получает или задает флаг заголовков листа.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
