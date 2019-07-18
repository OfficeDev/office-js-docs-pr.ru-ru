---
title: Набор обязательных элементов API JavaScript для Excel 1,7
description: Сведения о наборе требований ExcelApi 1,7
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c84d099982225bae11cb3deba8a0503da0695aed
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771990"
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

| Класс | Поля | Описание |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Представляет тип диаграммы. Дополнительные сведения см. в статье Excel. ChartType.|
||[id](/javascript/api/excel/excel.chart#id)|Уникальный идентификатор диаграммы. Только для чтения.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Указывает, следует ли отображать все кнопки полей в сводной диаграмме.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[вокруг](/javascript/api/excel/excel.chartareaformat#border)|Представляет формат границы области диаграммы, включающий цвет, lineStyle и толщину. Только для чтения.|
|[Чартареаформатдата](/javascript/api/excel/excel.chartareaformatdata)|[вокруг](/javascript/api/excel/excel.chartareaformatdata#border)|Представляет формат границы области диаграммы, включающий цвет, lineStyle и толщину. Только для чтения.|
|[Чартареаформатлоадоптионс](/javascript/api/excel/excel.chartareaformatloadoptions)|[вокруг](/javascript/api/excel/excel.chartareaformatloadoptions#border)|Представляет формат границы области диаграммы, включающий цвет, lineStyle и толщину.|
|[Чартареаформатупдатедата](/javascript/api/excel/excel.chartareaformatupdatedata)|[вокруг](/javascript/api/excel/excel.chartareaformatupdatedata#border)|Представляет формат границы области диаграммы, включающий цвет, lineStyle и толщину.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[GetItem (тип: " \| недопустимое значение" Category " \| " значение " \| Series", Группа?: "основной" \| "дополнительный")](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Возвращает указанную ось, определенную по типу и группе.|
||[GetItem (тип: Excel. Чартаксистипе, Group?: Excel. Чартаксисграуп)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Возвращает указанную ось, определенную по типу и группе.|
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
|[Чартаксисдата](/javascript/api/excel/excel.chartaxisdata)|[axisGroup](/javascript/api/excel/excel.chartaxisdata#axisgroup)|Представляет группу для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксисграуп. Только для чтения.|
||[Басетимеунит](/javascript/api/excel/excel.chartaxisdata#basetimeunit)|Возвращает или задает базовую единицу измерений для указанной оси категории.|
||[categoryType](/javascript/api/excel/excel.chartaxisdata#categorytype)|Возвращает или задает тип оси категории.|
||[Кустомдисплайунит](/javascript/api/excel/excel.chartaxisdata#customdisplayunit)|Представляет значение отображаемой единицы измерения настраиваемой оси.  Только для чтения. Чтобы задать это свойство, используйте метод SetCustomDisplayUnit(double).|
||[displayUnit](/javascript/api/excel/excel.chartaxisdata#displayunit)|Представляет отображаемую единицу измерения оси. Дополнительные сведения см. в статье Excel. Чартаксисдисплайунит.|
||[height](/javascript/api/excel/excel.chartaxisdata#height)|Представляет высоту оси диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
||[left](/javascript/api/excel/excel.chartaxisdata#left)|Представляет расстояние от левого края оси до левой стороны области диаграммы (в пунктах).  Значение null, если ось не отображается. Только для чтения.|
||[logBase](/javascript/api/excel/excel.chartaxisdata#logbase)|Представляет базу логарифма при использовании логарифмических шкал.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisdata#majortickmark)|Представляет тип основного деления для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистиккмарк.|
||[Мажортимеунитскале](/javascript/api/excel/excel.chartaxisdata#majortimeunitscale)|Возвращает или задает основное значение шкалы единиц измерений для оси категории, если для свойства CategoryType установлено значение TimeScale.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisdata#minortickmark)|Представляет тип дополнительного деления для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистиккмарк.|
||[Минортимеунитскале](/javascript/api/excel/excel.chartaxisdata#minortimeunitscale)|Возвращает или задает дополнительное значение шкалы единиц измерений для оси категории, если для свойства CategoryType установлено значение TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisdata#reverseplotorder)|Указывает, отображает ли Microsoft Excel точки данных от последней к первой.|
||[scaleType](/javascript/api/excel/excel.chartaxisdata#scaletype)|Представляет тип шкалы оси значений. Дополнительные сведения см. в статье Excel. Чартаксисскалетипе.|
||[Шовдисплайунитлабел](/javascript/api/excel/excel.chartaxisdata#showdisplayunitlabel)|Указывает, видна ли метка отображаемой единицы измерения оси.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisdata#ticklabelposition)|Представляет положение подписей делений на указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистикклабелпоситион.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisdata#ticklabelspacing)|Представляет количество категорий или рядов между подписями делений. Может иметь значение от 1 до 31 999 или пустую строку для автоматической настройки. Возвращаемое значение всегда является числом.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisdata#tickmarkspacing)|Представляет количество категорий или рядов между делениями.|
||[top](/javascript/api/excel/excel.chartaxisdata#top)|Представляет расстояние от верхнего края оси до верха области диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
||[type](/javascript/api/excel/excel.chartaxisdata#type)|Представляет тип оси. Дополнительные сведения см. в статье Excel. Чартаксистипе.|
||[visible](/javascript/api/excel/excel.chartaxisdata#visible)|Логическое значение, представляющее видимость оси.|
||[width](/javascript/api/excel/excel.chartaxisdata#width)|Представляет ширину оси диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
|[Чартаксислоадоптионс](/javascript/api/excel/excel.chartaxisloadoptions)|[axisGroup](/javascript/api/excel/excel.chartaxisloadoptions#axisgroup)|Представляет группу для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксисграуп. Только для чтения.|
||[Басетимеунит](/javascript/api/excel/excel.chartaxisloadoptions#basetimeunit)|Возвращает или задает базовую единицу измерений для указанной оси категории.|
||[categoryType](/javascript/api/excel/excel.chartaxisloadoptions#categorytype)|Возвращает или задает тип оси категории.|
||[ее пересекает другая](/javascript/api/excel/excel.chartaxisloadoptions#crosses)|[УСТАРЕВШИй; сохраняется для обеспечения обратной совместимости с существующими первыми решениями]. Взамен рекомендуется `Position` использовать.|
||[crossesAt](/javascript/api/excel/excel.chartaxisloadoptions#crossesat)|[УСТАРЕВШИй; сохраняется для обеспечения обратной совместимости с существующими первыми решениями]. Взамен рекомендуется `PositionAt` использовать.|
||[Кустомдисплайунит](/javascript/api/excel/excel.chartaxisloadoptions#customdisplayunit)|Представляет значение отображаемой единицы измерения настраиваемой оси.  Только для чтения. Чтобы задать это свойство, используйте метод SetCustomDisplayUnit(double).|
||[displayUnit](/javascript/api/excel/excel.chartaxisloadoptions#displayunit)|Представляет отображаемую единицу измерения оси. Дополнительные сведения см. в статье Excel. Чартаксисдисплайунит.|
||[height](/javascript/api/excel/excel.chartaxisloadoptions#height)|Представляет высоту оси диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
||[left](/javascript/api/excel/excel.chartaxisloadoptions#left)|Представляет расстояние от левого края оси до левой стороны области диаграммы (в пунктах).  Значение null, если ось не отображается. Только для чтения.|
||[logBase](/javascript/api/excel/excel.chartaxisloadoptions#logbase)|Представляет базу логарифма при использовании логарифмических шкал.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#majortickmark)|Представляет тип основного деления для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистиккмарк.|
||[Мажортимеунитскале](/javascript/api/excel/excel.chartaxisloadoptions#majortimeunitscale)|Возвращает или задает основное значение шкалы единиц измерений для оси категории, если для свойства CategoryType установлено значение TimeScale.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#minortickmark)|Представляет тип дополнительного деления для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистиккмарк.|
||[Минортимеунитскале](/javascript/api/excel/excel.chartaxisloadoptions#minortimeunitscale)|Возвращает или задает дополнительное значение шкалы единиц измерений для оси категории, если для свойства CategoryType установлено значение TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisloadoptions#reverseplotorder)|Указывает, отображает ли Microsoft Excel точки данных от последней к первой.|
||[scaleType](/javascript/api/excel/excel.chartaxisloadoptions#scaletype)|Представляет тип шкалы оси значений. Дополнительные сведения см. в статье Excel. Чартаксисскалетипе.|
||[Шовдисплайунитлабел](/javascript/api/excel/excel.chartaxisloadoptions#showdisplayunitlabel)|Указывает, видна ли метка отображаемой единицы измерения оси.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelposition)|Представляет положение подписей делений на указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистикклабелпоситион.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelspacing)|Представляет количество категорий или рядов между подписями делений. Может иметь значение от 1 до 31 999 или пустую строку для автоматической настройки. Возвращаемое значение всегда является числом.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisloadoptions#tickmarkspacing)|Представляет количество категорий или рядов между делениями.|
||[top](/javascript/api/excel/excel.chartaxisloadoptions#top)|Представляет расстояние от верхнего края оси до верха области диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
||[type](/javascript/api/excel/excel.chartaxisloadoptions#type)|Представляет тип оси. Дополнительные сведения см. в статье Excel. Чартаксистипе.|
||[visible](/javascript/api/excel/excel.chartaxisloadoptions#visible)|Логическое значение, представляющее видимость оси.|
||[width](/javascript/api/excel/excel.chartaxisloadoptions#width)|Представляет ширину оси диаграммы (в пунктах). Значение null, если ось не отображается. Только для чтения.|
|[Чартаксисупдатедата](/javascript/api/excel/excel.chartaxisupdatedata)|[Басетимеунит](/javascript/api/excel/excel.chartaxisupdatedata#basetimeunit)|Возвращает или задает базовую единицу измерений для указанной оси категории.|
||[categoryType](/javascript/api/excel/excel.chartaxisupdatedata#categorytype)|Возвращает или задает тип оси категории.|
||[displayUnit](/javascript/api/excel/excel.chartaxisupdatedata#displayunit)|Представляет отображаемую единицу измерения оси. Дополнительные сведения см. в статье Excel. Чартаксисдисплайунит.|
||[logBase](/javascript/api/excel/excel.chartaxisupdatedata#logbase)|Представляет базу логарифма при использовании логарифмических шкал.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#majortickmark)|Представляет тип основного деления для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистиккмарк.|
||[Мажортимеунитскале](/javascript/api/excel/excel.chartaxisupdatedata#majortimeunitscale)|Возвращает или задает основное значение шкалы единиц измерений для оси категории, если для свойства CategoryType установлено значение TimeScale.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#minortickmark)|Представляет тип дополнительного деления для указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистиккмарк.|
||[Минортимеунитскале](/javascript/api/excel/excel.chartaxisupdatedata#minortimeunitscale)|Возвращает или задает дополнительное значение шкалы единиц измерений для оси категории, если для свойства CategoryType установлено значение TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisupdatedata#reverseplotorder)|Указывает, отображает ли Microsoft Excel точки данных от последней к первой.|
||[scaleType](/javascript/api/excel/excel.chartaxisupdatedata#scaletype)|Представляет тип шкалы оси значений. Дополнительные сведения см. в статье Excel. Чартаксисскалетипе.|
||[Шовдисплайунитлабел](/javascript/api/excel/excel.chartaxisupdatedata#showdisplayunitlabel)|Указывает, видна ли метка отображаемой единицы измерения оси.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelposition)|Представляет положение подписей делений на указанной оси. Дополнительные сведения см. в статье Excel. Чартаксистикклабелпоситион.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelspacing)|Представляет количество категорий или рядов между подписями делений. Может иметь значение от 1 до 31 999 или пустую строку для автоматической настройки. Возвращаемое значение всегда является числом.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisupdatedata#tickmarkspacing)|Представляет количество категорий или рядов между делениями.|
||[visible](/javascript/api/excel/excel.chartaxisupdatedata#visible)|Логическое значение, представляющее видимость оси.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|HTML-код цвета, представляющий цвет границ в диаграмме.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Представляет тип линии границы. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[Set (Properties: Excel. ChartBorder)](/javascript/api/excel/excel.chartborder#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартбордерупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartborder#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Представляет толщину границы (в пунктах).|
|[Чартбордердата](/javascript/api/excel/excel.chartborderdata)|[color](/javascript/api/excel/excel.chartborderdata#color)|HTML-код цвета, представляющий цвет границ в диаграмме.|
||[lineStyle](/javascript/api/excel/excel.chartborderdata#linestyle)|Представляет тип линии границы. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartborderdata#weight)|Представляет толщину границы (в пунктах).|
|[Чартбордерлоадоптионс](/javascript/api/excel/excel.chartborderloadoptions)|[$all](/javascript/api/excel/excel.chartborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartborderloadoptions#color)|HTML-код цвета, представляющий цвет границ в диаграмме.|
||[lineStyle](/javascript/api/excel/excel.chartborderloadoptions#linestyle)|Представляет тип линии границы. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartborderloadoptions#weight)|Представляет толщину границы (в пунктах).|
|[Чартбордерупдатедата](/javascript/api/excel/excel.chartborderupdatedata)|[color](/javascript/api/excel/excel.chartborderupdatedata#color)|HTML-код цвета, представляющий цвет границ в диаграмме.|
||[lineStyle](/javascript/api/excel/excel.chartborderupdatedata#linestyle)|Представляет тип линии границы. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartborderupdatedata#weight)|Представляет толщину границы (в пунктах).|
|[Чартколлектионлоадоптионс](/javascript/api/excel/excel.chartcollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartcollectionloadoptions#charttype)|Для каждого элемента в коллекции: представляет тип диаграммы. Дополнительные сведения см. в статье Excel. ChartType.|
||[id](/javascript/api/excel/excel.chartcollectionloadoptions#id)|Для каждого элемента в коллекции: уникальный идентификатор диаграммы. Только для чтения.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartcollectionloadoptions#showallfieldbuttons)|Для каждого элемента в коллекции: указывает, нужно ли отображать все кнопки полей в сводной диаграмме.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[chartType](/javascript/api/excel/excel.chartdata#charttype)|Представляет тип диаграммы. Дополнительные сведения см. в статье Excel. ChartType.|
||[id](/javascript/api/excel/excel.chartdata#id)|Уникальный идентификатор диаграммы. Только для чтения.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartdata#showallfieldbuttons)|Указывает, следует ли отображать все кнопки полей в сводной диаграмме.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[символ](/javascript/api/excel/excel.chartdatalabel#separator)|Строка, представляющая разделитель для метки данных на диаграмме.|
||[Set (Properties: Excel. Чартдаталабел)](/javascript/api/excel/excel.chartdatalabel#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартдаталабелупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartdatalabel#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[Чартдаталабелдата](/javascript/api/excel/excel.chartdatalabeldata)|[position](/javascript/api/excel/excel.chartdatalabeldata#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[символ](/javascript/api/excel/excel.chartdatalabeldata#separator)|Строка, представляющая разделитель для метки данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabeldata#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabeldata#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabeldata#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabeldata#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabeldata#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabeldata#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[Чартдаталабеллоадоптионс](/javascript/api/excel/excel.chartdatalabelloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelloadoptions#$all)||
||[position](/javascript/api/excel/excel.chartdatalabelloadoptions#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[символ](/javascript/api/excel/excel.chartdatalabelloadoptions#separator)|Строка, представляющая разделитель для метки данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelloadoptions#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelloadoptions#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelloadoptions#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelloadoptions#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelloadoptions#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabelloadoptions#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[Чартдаталабелупдатедата](/javascript/api/excel/excel.chartdatalabelupdatedata)|[position](/javascript/api/excel/excel.chartdatalabelupdatedata#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[символ](/javascript/api/excel/excel.chartdatalabelupdatedata#separator)|Строка, представляющая разделитель для метки данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelupdatedata#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelupdatedata#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelupdatedata#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelupdatedata#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelupdatedata#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabelupdatedata#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[Чартформатстринг](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. для объекта "символы диаграммы".|
||[Set (Properties: Excel. Чартформатстринг)](/javascript/api/excel/excel.chartformatstring#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартформатстрингупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartformatstring#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартформатстрингдата](/javascript/api/excel/excel.chartformatstringdata)|[font](/javascript/api/excel/excel.chartformatstringdata#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. для объекта "символы диаграммы".|
|[Чартформатстринглоадоптионс](/javascript/api/excel/excel.chartformatstringloadoptions)|[$all](/javascript/api/excel/excel.chartformatstringloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartformatstringloadoptions#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. для объекта "символы диаграммы".|
|[Чартформатстрингупдатедата](/javascript/api/excel/excel.chartformatstringupdatedata)|[font](/javascript/api/excel/excel.chartformatstringupdatedata#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. для объекта "символы диаграммы".|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Представляет высоту условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Представляет левую (в пунктах) условные обозначения диаграммы. Значение null, если условные обозначения не отображаются.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Представляет коллекцию объектов legendEntries в условных обозначениях. Только для чтения.|
||[Шовшадов](/javascript/api/excel/excel.chartlegend#showshadow)|Указывает, имеет ли легенда тень на диаграмме.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Представляет верх условных обозначений диаграммы.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Представляет ширину условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
|[Чартлеженддата](/javascript/api/excel/excel.chartlegenddata)|[height](/javascript/api/excel/excel.chartlegenddata#height)|Представляет высоту условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
||[left](/javascript/api/excel/excel.chartlegenddata#left)|Представляет левую (в пунктах) условные обозначения диаграммы. Значение null, если условные обозначения не отображаются.|
||[legendEntries](/javascript/api/excel/excel.chartlegenddata#legendentries)|Представляет коллекцию объектов legendEntries в условных обозначениях. Только для чтения.|
||[Шовшадов](/javascript/api/excel/excel.chartlegenddata#showshadow)|Указывает, имеет ли легенда тень на диаграмме.|
||[top](/javascript/api/excel/excel.chartlegenddata#top)|Представляет верх условных обозначений диаграммы.|
||[width](/javascript/api/excel/excel.chartlegenddata#width)|Представляет ширину условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
|[Чартлежендентри](/javascript/api/excel/excel.chartlegendentry)|[Set (Properties: Excel. Чартлежендентри)](/javascript/api/excel/excel.chartlegendentry#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартлежендентрюпдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartlegendentry#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Представляет видимый элемент записи условных обозначений диаграммы.|
|[Чартлежендентриколлектион](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Возвращает количество legendEntry в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Возвращает объект legendEntry по указанному индексу.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Чартлежендентриколлектионлоадоптионс](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#visible)|Для каждого элемента в коллекции — представляет видимую запись легенды диаграммы.|
|[Чартлежендентридата](/javascript/api/excel/excel.chartlegendentrydata)|[visible](/javascript/api/excel/excel.chartlegendentrydata#visible)|Представляет видимый элемент записи условных обозначений диаграммы.|
|[Чартлежендентрилоадоптионс](/javascript/api/excel/excel.chartlegendentryloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentryloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentryloadoptions#visible)|Представляет видимый элемент записи условных обозначений диаграммы.|
|[Чартлежендентрюпдатедата](/javascript/api/excel/excel.chartlegendentryupdatedata)|[visible](/javascript/api/excel/excel.chartlegendentryupdatedata#visible)|Представляет видимый элемент записи условных обозначений диаграммы.|
|[Чартлежендлоадоптионс](/javascript/api/excel/excel.chartlegendloadoptions)|[height](/javascript/api/excel/excel.chartlegendloadoptions#height)|Представляет высоту условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
||[left](/javascript/api/excel/excel.chartlegendloadoptions#left)|Представляет левую (в пунктах) условные обозначения диаграммы. Значение null, если условные обозначения не отображаются.|
||[Шовшадов](/javascript/api/excel/excel.chartlegendloadoptions#showshadow)|Указывает, имеет ли легенда тень на диаграмме.|
||[top](/javascript/api/excel/excel.chartlegendloadoptions#top)|Представляет верх условных обозначений диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendloadoptions#width)|Представляет ширину условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
|[Чартлежендупдатедата](/javascript/api/excel/excel.chartlegendupdatedata)|[height](/javascript/api/excel/excel.chartlegendupdatedata#height)|Представляет высоту условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
||[left](/javascript/api/excel/excel.chartlegendupdatedata#left)|Представляет левую (в пунктах) условные обозначения диаграммы. Значение null, если условные обозначения не отображаются.|
||[Шовшадов](/javascript/api/excel/excel.chartlegendupdatedata#showshadow)|Указывает, имеет ли легенда тень на диаграмме.|
||[top](/javascript/api/excel/excel.chartlegendupdatedata#top)|Представляет верх условных обозначений диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendupdatedata#width)|Представляет ширину условных обозначений на диаграмме в пунктах. Значение null, если условные обозначения не отображаются.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Представляет стиль линии. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Представляет толщину линии (в пунктах).|
|[Чартлинеформатдата](/javascript/api/excel/excel.chartlineformatdata)|[lineStyle](/javascript/api/excel/excel.chartlineformatdata#linestyle)|Представляет стиль линии. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartlineformatdata#weight)|Представляет толщину линии (в пунктах).|
|[Чартлинеформатлоадоптионс](/javascript/api/excel/excel.chartlineformatloadoptions)|[lineStyle](/javascript/api/excel/excel.chartlineformatloadoptions#linestyle)|Представляет стиль линии. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartlineformatloadoptions#weight)|Представляет толщину линии (в пунктах).|
|[Чартлинеформатупдатедата](/javascript/api/excel/excel.chartlineformatupdatedata)|[lineStyle](/javascript/api/excel/excel.chartlineformatupdatedata#linestyle)|Представляет стиль линии. Дополнительные сведения см. в статье Excel. Чартлинестиле.|
||[weight](/javascript/api/excel/excel.chartlineformatupdatedata#weight)|Представляет толщину линии (в пунктах).|
|[Чартлоадоптионс](/javascript/api/excel/excel.chartloadoptions)|[chartType](/javascript/api/excel/excel.chartloadoptions#charttype)|Представляет тип диаграммы. Дополнительные сведения см. в статье Excel. ChartType.|
||[id](/javascript/api/excel/excel.chartloadoptions#id)|Уникальный идентификатор диаграммы. Только для чтения.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartloadoptions#showallfieldbuttons)|Указывает, следует ли отображать все кнопки полей в сводной диаграмме.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Указывает, имеет ли точка данных метку данных. Неприменимо для поверхностных диаграмм.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Представление цветового HTML-кода для цвета фона маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Представление цветового HTML-кода для цвета переднего плана маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Представляет стиль маркера точки данных диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Возвращает метку данных точки диаграммы. Только для чтения.|
|[Чартпоинтдата](/javascript/api/excel/excel.chartpointdata)|[dataLabel](/javascript/api/excel/excel.chartpointdata#datalabel)|Возвращает метку данных точки диаграммы. Только для чтения.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointdata#hasdatalabel)|Указывает, имеет ли точка данных метку данных. Неприменимо для поверхностных диаграмм.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointdata#markerbackgroundcolor)|Представление цветового HTML-кода для цвета фона маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointdata#markerforegroundcolor)|Представление цветового HTML-кода для цвета переднего плана маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerSize](/javascript/api/excel/excel.chartpointdata#markersize)|Представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpointdata#markerstyle)|Представляет стиль маркера точки данных диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[вокруг](/javascript/api/excel/excel.chartpointformat#border)|Представляет формат границы точки данных диаграммы, включающий сведения о цвете, стиле и весу. Только для чтения.|
|[Чартпоинтформатдата](/javascript/api/excel/excel.chartpointformatdata)|[вокруг](/javascript/api/excel/excel.chartpointformatdata#border)|Представляет формат границы точки данных диаграммы, включающий сведения о цвете, стиле и весу. Только для чтения.|
|[Чартпоинтформатлоадоптионс](/javascript/api/excel/excel.chartpointformatloadoptions)|[вокруг](/javascript/api/excel/excel.chartpointformatloadoptions#border)|Представляет формат границы точки данных диаграммы, включающий сведения о цвете, стиле и весу.|
|[Чартпоинтформатупдатедата](/javascript/api/excel/excel.chartpointformatupdatedata)|[вокруг](/javascript/api/excel/excel.chartpointformatupdatedata#border)|Представляет формат границы точки данных диаграммы, включающий сведения о цвете, стиле и весу.|
|[Чартпоинтлоадоптионс](/javascript/api/excel/excel.chartpointloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointloadoptions#datalabel)|Возвращает метку данных точки диаграммы.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointloadoptions#hasdatalabel)|Указывает, имеет ли точка данных метку данных. Неприменимо для поверхностных диаграмм.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerbackgroundcolor)|Представление цветового HTML-кода для цвета фона маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerforegroundcolor)|Представление цветового HTML-кода для цвета переднего плана маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerSize](/javascript/api/excel/excel.chartpointloadoptions#markersize)|Представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpointloadoptions#markerstyle)|Представляет стиль маркера точки данных диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
|[Чартпоинтупдатедата](/javascript/api/excel/excel.chartpointupdatedata)|[dataLabel](/javascript/api/excel/excel.chartpointupdatedata#datalabel)|Возвращает метку данных точки диаграммы.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointupdatedata#hasdatalabel)|Указывает, имеет ли точка данных метку данных. Неприменимо для поверхностных диаграмм.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerbackgroundcolor)|Представление цветового HTML-кода для цвета фона маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerforegroundcolor)|Представление цветового HTML-кода для цвета переднего плана маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerSize](/javascript/api/excel/excel.chartpointupdatedata#markersize)|Представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpointupdatedata#markerstyle)|Представляет стиль маркера точки данных диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
|[Чартпоинтсколлектионлоадоптионс](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#datalabel)|Для каждого элемента в коллекции: Возвращает метку данных точки диаграммы.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#hasdatalabel)|Для каждого элемента в коллекции: указывает, имеет ли точка данных метку данных. Неприменимо для поверхностных диаграмм.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerbackgroundcolor)|Для каждого элемента в коллекции: цветовое HTML-представление цвета фона маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerforegroundcolor)|Для каждого элемента в коллекции: цвет HTML-кода представления цвета текста маркера точки данных. Например, #FF0000 обозначает красный.|
||[markerSize](/javascript/api/excel/excel.chartpointscollectionloadoptions#markersize)|Для каждого элемента в коллекции: представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerstyle)|Для каждого элемента в коллекции: представляет стиль маркера точки данных диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
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
|[Чартсериесколлектионлоадоптионс](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartseriescollectionloadoptions#charttype)|Для каждого элемента в коллекции: представляет тип диаграммы ряда. Дополнительные сведения см. в статье Excel. ChartType.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#doughnutholesize)|Для каждого элемента в коллекции: представляет размер отверстия для ряда диаграммы.  Допустимо только в doughnutExploded и кольцевых диаграммах.|
||[отсортирован](/javascript/api/excel/excel.chartseriescollectionloadoptions#filtered)|Для каждого элемента в коллекции: логическое значение, указывающее, отфильтрована ли серия. Неприменимо для поверхностных диаграмм.|
||[gapWidth](/javascript/api/excel/excel.chartseriescollectionloadoptions#gapwidth)|Для каждого элемента в коллекции: представляет ширину зазора для ряда диаграммы.  Допустимо только для линейчатых диаграмм и гистограмм, а также|
||[hasDataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#hasdatalabels)|Для каждого элемента в коллекции: логическое значение, указывающее, есть ли в рядах метки данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerbackgroundcolor)|Для каждого элемента в коллекции: представляет цвет фона маркеров ряда диаграммы.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerforegroundcolor)|Для каждого элемента в коллекции: представляет основной цвет маркеров ряда диаграммы.|
||[markerSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#markersize)|Для каждого элемента в коллекции: представляет размер маркера для ряда диаграммы.|
||[markerStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerstyle)|Для каждого элемента в коллекции: представляет стиль маркера ряда диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
||[plotOrder](/javascript/api/excel/excel.chartseriescollectionloadoptions#plotorder)|Для каждого элемента в коллекции: представляет порядок построения рядов диаграммы в группе диаграммы.|
||[Шовшадов](/javascript/api/excel/excel.chartseriescollectionloadoptions#showshadow)|Для каждого элемента в коллекции: логическое значение, представляющее, имеет ли ряд теневую копию.|
||[высокое](/javascript/api/excel/excel.chartseriescollectionloadoptions#smooth)|Для каждого элемента в коллекции: логическое значение, указывающее, является ли ряд гладким или нет. Применяется только к графикам и точечным диаграммам.|
|[Чартсериесдата](/javascript/api/excel/excel.chartseriesdata)|[chartType](/javascript/api/excel/excel.chartseriesdata#charttype)|Представляет тип диаграммы для ряда. Дополнительные сведения см. в статье Excel. ChartType.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesdata#doughnutholesize)|Представляет размер отверстия ряда кольцевой диаграммы.  Допустимо только в doughnutExploded и кольцевых диаграммах.|
||[отсортирован](/javascript/api/excel/excel.chartseriesdata#filtered)|Логическое значение, которое указывает, фильтруется ли ряд. Неприменимо для поверхностных диаграмм.|
||[gapWidth](/javascript/api/excel/excel.chartseriesdata#gapwidth)|Представляет ширину разрывов рядов диаграммы.  Допустимо только для линейчатых диаграмм и гистограмм, а также|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesdata#hasdatalabels)|Логическое значение, которое указывает, имеют ли ряды метки данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesdata#markerbackgroundcolor)|Представляет цвет фона маркеров для рядов диаграммы.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesdata#markerforegroundcolor)|Представляет цвет переднего плана для рядов диаграммы.|
||[markerSize](/javascript/api/excel/excel.chartseriesdata#markersize)|Представляет размер маркера рядов диаграммы.|
||[markerStyle](/javascript/api/excel/excel.chartseriesdata#markerstyle)|Представляет стиль маркера рядов диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
||[plotOrder](/javascript/api/excel/excel.chartseriesdata#plotorder)|Представляет порядок построения рядов диаграммы в группе диаграммы.|
||[Шовшадов](/javascript/api/excel/excel.chartseriesdata#showshadow)|Логическое значение, указывающее, есть ли у ряда теневая копия.|
||[высокое](/javascript/api/excel/excel.chartseriesdata#smooth)|Логическое значение, которое указывает, является ли ряд плавным.  Применяется только к графикам и точечным диаграммам.|
||[trendlines](/javascript/api/excel/excel.chartseriesdata#trendlines)|Представляет коллекцию линий тренда в ряду. Только для чтения.|
|[Чартсериеслоадоптионс](/javascript/api/excel/excel.chartseriesloadoptions)|[chartType](/javascript/api/excel/excel.chartseriesloadoptions#charttype)|Представляет тип диаграммы для ряда. Дополнительные сведения см. в статье Excel. ChartType.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesloadoptions#doughnutholesize)|Представляет размер отверстия ряда кольцевой диаграммы.  Допустимо только в doughnutExploded и кольцевых диаграммах.|
||[отсортирован](/javascript/api/excel/excel.chartseriesloadoptions#filtered)|Логическое значение, которое указывает, фильтруется ли ряд. Неприменимо для поверхностных диаграмм.|
||[gapWidth](/javascript/api/excel/excel.chartseriesloadoptions#gapwidth)|Представляет ширину разрывов рядов диаграммы.  Допустимо только для линейчатых диаграмм и гистограмм, а также|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesloadoptions#hasdatalabels)|Логическое значение, которое указывает, имеют ли ряды метки данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerbackgroundcolor)|Представляет цвет фона маркеров для рядов диаграммы.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerforegroundcolor)|Представляет цвет переднего плана для рядов диаграммы.|
||[markerSize](/javascript/api/excel/excel.chartseriesloadoptions#markersize)|Представляет размер маркера рядов диаграммы.|
||[markerStyle](/javascript/api/excel/excel.chartseriesloadoptions#markerstyle)|Представляет стиль маркера рядов диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
||[plotOrder](/javascript/api/excel/excel.chartseriesloadoptions#plotorder)|Представляет порядок построения рядов диаграммы в группе диаграммы.|
||[Шовшадов](/javascript/api/excel/excel.chartseriesloadoptions#showshadow)|Логическое значение, указывающее, есть ли у ряда теневая копия.|
||[высокое](/javascript/api/excel/excel.chartseriesloadoptions#smooth)|Логическое значение, которое указывает, является ли ряд плавным.  Применяется только к графикам и точечным диаграммам.|
|[Чартсериесупдатедата](/javascript/api/excel/excel.chartseriesupdatedata)|[chartType](/javascript/api/excel/excel.chartseriesupdatedata#charttype)|Представляет тип диаграммы для ряда. Дополнительные сведения см. в статье Excel. ChartType.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesupdatedata#doughnutholesize)|Представляет размер отверстия ряда кольцевой диаграммы.  Допустимо только в doughnutExploded и кольцевых диаграммах.|
||[отсортирован](/javascript/api/excel/excel.chartseriesupdatedata#filtered)|Логическое значение, которое указывает, фильтруется ли ряд. Неприменимо для поверхностных диаграмм.|
||[gapWidth](/javascript/api/excel/excel.chartseriesupdatedata#gapwidth)|Представляет ширину разрывов рядов диаграммы.  Допустимо только для линейчатых диаграмм и гистограмм, а также|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesupdatedata#hasdatalabels)|Логическое значение, которое указывает, имеют ли ряды метки данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerbackgroundcolor)|Представляет цвет фона маркеров для рядов диаграммы.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerforegroundcolor)|Представляет цвет переднего плана для рядов диаграммы.|
||[markerSize](/javascript/api/excel/excel.chartseriesupdatedata#markersize)|Представляет размер маркера рядов диаграммы.|
||[markerStyle](/javascript/api/excel/excel.chartseriesupdatedata#markerstyle)|Представляет стиль маркера рядов диаграммы. Дополнительные сведения см. в статье Excel. Чартмаркерстиле.|
||[plotOrder](/javascript/api/excel/excel.chartseriesupdatedata#plotorder)|Представляет порядок построения рядов диаграммы в группе диаграммы.|
||[Шовшадов](/javascript/api/excel/excel.chartseriesupdatedata#showshadow)|Логическое значение, указывающее, есть ли у ряда теневая копия.|
||[высокое](/javascript/api/excel/excel.chartseriesupdatedata#smooth)|Логическое значение, которое указывает, является ли ряд плавным.  Применяется только к графикам и точечным диаграммам.|
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
|[Чарттитледата](/javascript/api/excel/excel.charttitledata)|[height](/javascript/api/excel/excel.charttitledata#height)|Возвращает высоту заголовка диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается. Только для чтения.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitledata#horizontalalignment)|Представляет горизонтальное выравнивание для заголовка диаграммы.|
||[left](/javascript/api/excel/excel.charttitledata#left)|Представляет расстояние от левого края заголовка диаграммы до левого края области диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается.|
||[position](/javascript/api/excel/excel.charttitledata#position)|Представляет положение заголовка диаграммы. Дополнительные сведения см. в статье Excel. Чарттитлепоситион.|
||[Шовшадов](/javascript/api/excel/excel.charttitledata#showshadow)|Представляет логическое значение, которое определяет, имеет ли заголовок диаграммы тень.|
||[textOrientation](/javascript/api/excel/excel.charttitledata#textorientation)|Представляет ориентацию текста для заголовка диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.charttitledata#top)|Представляет расстояние от верхнего края заголовка диаграммы до верха области диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.charttitledata#verticalalignment)|Представляет вертикальное выравнивание для заголовка диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
||[width](/javascript/api/excel/excel.charttitledata#width)|Возвращает ширину заголовка диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается. Только для чтения.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[вокруг](/javascript/api/excel/excel.charttitleformat#border)|Представляет формат границы заголовка диаграммы, включающий цвет, lineStyle и толщину. Только для чтения.|
|[Чарттитлеформатдата](/javascript/api/excel/excel.charttitleformatdata)|[вокруг](/javascript/api/excel/excel.charttitleformatdata#border)|Представляет формат границы заголовка диаграммы, включающий цвет, lineStyle и толщину. Только для чтения.|
|[Чарттитлеформатлоадоптионс](/javascript/api/excel/excel.charttitleformatloadoptions)|[вокруг](/javascript/api/excel/excel.charttitleformatloadoptions#border)|Представляет формат границы заголовка диаграммы, включающий цвет, lineStyle и толщину.|
|[Чарттитлеформатупдатедата](/javascript/api/excel/excel.charttitleformatupdatedata)|[вокруг](/javascript/api/excel/excel.charttitleformatupdatedata#border)|Представляет формат границы заголовка диаграммы, включающий цвет, lineStyle и толщину.|
|[Чарттитлелоадоптионс](/javascript/api/excel/excel.charttitleloadoptions)|[height](/javascript/api/excel/excel.charttitleloadoptions#height)|Возвращает высоту заголовка диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается. Только для чтения.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitleloadoptions#horizontalalignment)|Представляет горизонтальное выравнивание для заголовка диаграммы.|
||[left](/javascript/api/excel/excel.charttitleloadoptions#left)|Представляет расстояние от левого края заголовка диаграммы до левого края области диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается.|
||[position](/javascript/api/excel/excel.charttitleloadoptions#position)|Представляет положение заголовка диаграммы. Дополнительные сведения см. в статье Excel. Чарттитлепоситион.|
||[Шовшадов](/javascript/api/excel/excel.charttitleloadoptions#showshadow)|Представляет логическое значение, которое определяет, имеет ли заголовок диаграммы тень.|
||[textOrientation](/javascript/api/excel/excel.charttitleloadoptions#textorientation)|Представляет ориентацию текста для заголовка диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.charttitleloadoptions#top)|Представляет расстояние от верхнего края заголовка диаграммы до верха области диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.charttitleloadoptions#verticalalignment)|Представляет вертикальное выравнивание для заголовка диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
||[width](/javascript/api/excel/excel.charttitleloadoptions#width)|Возвращает ширину заголовка диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается. Только для чтения.|
|[Чарттитлеупдатедата](/javascript/api/excel/excel.charttitleupdatedata)|[horizontalAlignment](/javascript/api/excel/excel.charttitleupdatedata#horizontalalignment)|Представляет горизонтальное выравнивание для заголовка диаграммы.|
||[left](/javascript/api/excel/excel.charttitleupdatedata#left)|Представляет расстояние от левого края заголовка диаграммы до левого края области диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается.|
||[position](/javascript/api/excel/excel.charttitleupdatedata#position)|Представляет положение заголовка диаграммы. Дополнительные сведения см. в статье Excel. Чарттитлепоситион.|
||[Шовшадов](/javascript/api/excel/excel.charttitleupdatedata#showshadow)|Представляет логическое значение, которое определяет, имеет ли заголовок диаграммы тень.|
||[textOrientation](/javascript/api/excel/excel.charttitleupdatedata#textorientation)|Представляет ориентацию текста для заголовка диаграммы. Значение должно быть целым числом от -90 до 90 или 180 для вертикально-ориентированного текста.|
||[top](/javascript/api/excel/excel.charttitleupdatedata#top)|Представляет расстояние от верхнего края заголовка диаграммы до верха области диаграммы (в пунктах). Значение null, если заголовок диаграммы не отображается.|
||[verticalAlignment](/javascript/api/excel/excel.charttitleupdatedata#verticalalignment)|Представляет вертикальное выравнивание для заголовка диаграммы. Дополнительные сведения см. в статье Excel. Чарттекствертикалалигнмент.|
|[Чарттрендлине](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Удаляет объект линии тренда.|
||[SBM](/javascript/api/excel/excel.charttrendline#intercept)|Представляет значение отсекаемого отрезка линии тренда. Можно указать в виде числового значения или пустой строки (для автоматически заданных значений). Возвращаемое значение всегда является числом.|
||[Мовингаверажепериод](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Представляет период линии тренда диаграммы. Применяется только для линии тренда с типом MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Представляет имя линии тренда. Можно указать в виде строкового значения или присвоить значение NULL для автоматических значений. Возвращаемое значение всегда является строковым|
||[Полиномиалордер](/javascript/api/excel/excel.charttrendline#polynomialorder)|Представляет порядок линии тренда диаграммы. Применяется только для линии тренда с типом полинома.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Представляет форматирование линии тренда диаграммы.|
||[Set (Properties: Excel. Чарттрендлине)](/javascript/api/excel/excel.charttrendline#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чарттрендлинеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.charttrendline#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Представляет тип линии тренда диаграммы.|
|[Чарттрендлинеколлектион](/javascript/api/excel/excel.charttrendlinecollection)|[Добавить (тип?: "линейный \| " экспоненциальный \| "" логарифмическое \| "" " \| MovingAverage" \| "" степень "")](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Добавляет новую линию тренда в коллекцию линий тренда.|
||[Add (Type?: Excel. Чарттрендлинетипе)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Добавляет новую линию тренда в коллекцию линий тренда.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Возвращает количество линий тренда в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Получает объект линии тренда по индексу, который является порядком вставки в массиве элементов.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Чарттрендлинеколлектионлоадоптионс](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#format)|Для каждого элемента в коллекции: представляет форматирование линии тренда диаграммы.|
||[SBM](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#intercept)|Для каждого элемента в коллекции — представляет значение параметра "конст" линии тренда. Можно указать в виде числового значения или пустой строки (для автоматически заданных значений). Возвращаемое значение всегда является числом.|
||[Мовингаверажепериод](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#movingaverageperiod)|Для каждого элемента в коллекции: представляет период линии тренда диаграммы. Применяется только для линии тренда с типом MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#name)|Для каждого элемента в коллекции: представляет название линии тренда. Можно указать в виде строкового значения или присвоить значение NULL для автоматических значений. Возвращаемое значение всегда является строковым|
||[Полиномиалордер](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#polynomialorder)|Для каждого элемента в коллекции: представляет порядок линии тренда диаграммы. Применяется только для линии тренда с типом полинома.|
||[type](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#type)|Для каждого элемента в коллекции: представляет тип линии тренда диаграммы.|
|[Чарттрендлинедата](/javascript/api/excel/excel.charttrendlinedata)|[format](/javascript/api/excel/excel.charttrendlinedata#format)|Представляет форматирование линии тренда диаграммы.|
||[SBM](/javascript/api/excel/excel.charttrendlinedata#intercept)|Представляет значение отсекаемого отрезка линии тренда. Можно указать в виде числового значения или пустой строки (для автоматически заданных значений). Возвращаемое значение всегда является числом.|
||[Мовингаверажепериод](/javascript/api/excel/excel.charttrendlinedata#movingaverageperiod)|Представляет период линии тренда диаграммы. Применяется только для линии тренда с типом MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlinedata#name)|Представляет имя линии тренда. Можно указать в виде строкового значения или присвоить значение NULL для автоматических значений. Возвращаемое значение всегда является строковым|
||[Полиномиалордер](/javascript/api/excel/excel.charttrendlinedata#polynomialorder)|Представляет порядок линии тренда диаграммы. Применяется только для линии тренда с типом полинома.|
||[type](/javascript/api/excel/excel.charttrendlinedata#type)|Представляет тип линии тренда диаграммы.|
|[Чарттрендлинеформат](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Представляет форматирование линий диаграммы. Только для чтения.|
||[Set (Properties: Excel. Чарттрендлинеформат)](/javascript/api/excel/excel.charttrendlineformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чарттрендлинеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.charttrendlineformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чарттрендлинеформатдата](/javascript/api/excel/excel.charttrendlineformatdata)|[line](/javascript/api/excel/excel.charttrendlineformatdata#line)|Представляет форматирование линий диаграммы. Только для чтения.|
|[Чарттрендлинеформатлоадоптионс](/javascript/api/excel/excel.charttrendlineformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charttrendlineformatloadoptions#line)|Представляет форматирование линий диаграммы.|
|[Чарттрендлинеформатупдатедата](/javascript/api/excel/excel.charttrendlineformatupdatedata)|[line](/javascript/api/excel/excel.charttrendlineformatupdatedata#line)|Представляет форматирование линий диаграммы.|
|[Чарттрендлинелоадоптионс](/javascript/api/excel/excel.charttrendlineloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlineloadoptions#format)|Представляет форматирование линии тренда диаграммы.|
||[SBM](/javascript/api/excel/excel.charttrendlineloadoptions#intercept)|Представляет значение отсекаемого отрезка линии тренда. Можно указать в виде числового значения или пустой строки (для автоматически заданных значений). Возвращаемое значение всегда является числом.|
||[Мовингаверажепериод](/javascript/api/excel/excel.charttrendlineloadoptions#movingaverageperiod)|Представляет период линии тренда диаграммы. Применяется только для линии тренда с типом MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlineloadoptions#name)|Представляет имя линии тренда. Можно указать в виде строкового значения или присвоить значение NULL для автоматических значений. Возвращаемое значение всегда является строковым|
||[Полиномиалордер](/javascript/api/excel/excel.charttrendlineloadoptions#polynomialorder)|Представляет порядок линии тренда диаграммы. Применяется только для линии тренда с типом полинома.|
||[type](/javascript/api/excel/excel.charttrendlineloadoptions#type)|Представляет тип линии тренда диаграммы.|
|[Чарттрендлинеупдатедата](/javascript/api/excel/excel.charttrendlineupdatedata)|[format](/javascript/api/excel/excel.charttrendlineupdatedata#format)|Представляет форматирование линии тренда диаграммы.|
||[SBM](/javascript/api/excel/excel.charttrendlineupdatedata#intercept)|Представляет значение отсекаемого отрезка линии тренда. Можно указать в виде числового значения или пустой строки (для автоматически заданных значений). Возвращаемое значение всегда является числом.|
||[Мовингаверажепериод](/javascript/api/excel/excel.charttrendlineupdatedata#movingaverageperiod)|Представляет период линии тренда диаграммы. Применяется только для линии тренда с типом MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlineupdatedata#name)|Представляет имя линии тренда. Можно указать в виде строкового значения или присвоить значение NULL для автоматических значений. Возвращаемое значение всегда является строковым|
||[Полиномиалордер](/javascript/api/excel/excel.charttrendlineupdatedata#polynomialorder)|Представляет порядок линии тренда диаграммы. Применяется только для линии тренда с типом полинома.|
||[type](/javascript/api/excel/excel.charttrendlineupdatedata#type)|Представляет тип линии тренда диаграммы.|
|[Чартупдатедата](/javascript/api/excel/excel.chartupdatedata)|[chartType](/javascript/api/excel/excel.chartupdatedata#charttype)|Представляет тип диаграммы. Дополнительные сведения см. в статье Excel. ChartType.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartupdatedata#showallfieldbuttons)|Указывает, следует ли отображать все кнопки полей в сводной диаграмме.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.customproperty#key)|Возвращает ключ настраиваемого свойства. Только для чтения.|
||[type](/javascript/api/excel/excel.customproperty#type)|Получает тип значения настраиваемого свойства. Только для чтения.|
||[Set (Properties: Excel. CustomProperty)](/javascript/api/excel/excel.customproperty#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кустомпропертюпдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.customproperty#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[value](/javascript/api/excel/excel.customproperty#value)|Получает или задает значение настраиваемого свойства.|
|[Кустомпропертиколлектион](/javascript/api/excel/excel.custompropertycollection)|[Add (Key: строка, Value: Any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Создает или задает настраиваемое свойство.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Вызывается, если настраиваемое свойство не существует.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Возвращает нулевой объект, если настраиваемое свойство не существует.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Кустомпропертиколлектионлоадоптионс](/javascript/api/excel/excel.custompropertycollectionloadoptions)|[$all](/javascript/api/excel/excel.custompropertycollectionloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertycollectionloadoptions#key)|Для каждого элемента в коллекции: получает ключ настраиваемого свойства. Только для чтения.|
||[type](/javascript/api/excel/excel.custompropertycollectionloadoptions#type)|Для каждого элемента в коллекции: получает тип значения настраиваемого свойства. Только для чтения.|
||[value](/javascript/api/excel/excel.custompropertycollectionloadoptions#value)|Для каждого элемента в коллекции: Получает или задает значение настраиваемого свойства.|
|[Кустомпропертидата](/javascript/api/excel/excel.custompropertydata)|[key](/javascript/api/excel/excel.custompropertydata#key)|Возвращает ключ настраиваемого свойства. Только для чтения.|
||[type](/javascript/api/excel/excel.custompropertydata#type)|Получает тип значения настраиваемого свойства. Только для чтения.|
||[value](/javascript/api/excel/excel.custompropertydata#value)|Получает или задает значение настраиваемого свойства.|
|[Кустомпропертилоадоптионс](/javascript/api/excel/excel.custompropertyloadoptions)|[$all](/javascript/api/excel/excel.custompropertyloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertyloadoptions#key)|Возвращает ключ настраиваемого свойства. Только для чтения.|
||[type](/javascript/api/excel/excel.custompropertyloadoptions#type)|Получает тип значения настраиваемого свойства. Только для чтения.|
||[value](/javascript/api/excel/excel.custompropertyloadoptions#value)|Получает или задает значение настраиваемого свойства.|
|[Кустомпропертюпдатедата](/javascript/api/excel/excel.custompropertyupdatedata)|[value](/javascript/api/excel/excel.custompropertyupdatedata#value)|Получает или задает значение настраиваемого свойства.|
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
||[Set (Properties: Excel. DocumentProperties)](/javascript/api/excel/excel.documentproperties#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Документпропертиесупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.documentproperties#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Получает или задает тему книги.|
||[заголовок](/javascript/api/excel/excel.documentproperties#title)|Получает или задает название книги.|
|[Документпропертиесдата](/javascript/api/excel/excel.documentpropertiesdata)|[Редактирование](/javascript/api/excel/excel.documentpropertiesdata#author)|Получает или задает автора книги.|
||[категории](/javascript/api/excel/excel.documentpropertiesdata#category)|Получает или задает категорию книги.|
||[comments](/javascript/api/excel/excel.documentpropertiesdata#comments)|Получает или задает примечания книги.|
||[company](/javascript/api/excel/excel.documentpropertiesdata#company)|Получает или задает компанию книги.|
||[creationDate](/javascript/api/excel/excel.documentpropertiesdata#creationdate)|Получает дату создания книги. Только для чтения.|
||[собственный](/javascript/api/excel/excel.documentpropertiesdata#custom)|Получает коллекцию настраиваемых свойств книги. Только для чтения.|
||[keyword](/javascript/api/excel/excel.documentpropertiesdata#keywords)|Получает или задает ключевые слова книги.|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesdata#lastauthor)|Получает последнего автора книги. Только для чтения.|
||[manager](/javascript/api/excel/excel.documentpropertiesdata#manager)|Получает или задает менеджера книги.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesdata#revisionnumber)|Получает номер редакции книги. Только для чтения.|
||[subject](/javascript/api/excel/excel.documentpropertiesdata#subject)|Получает или задает тему книги.|
||[заголовок](/javascript/api/excel/excel.documentpropertiesdata#title)|Получает или задает название книги.|
|[Документпропертиеслоадоптионс](/javascript/api/excel/excel.documentpropertiesloadoptions)|[$all](/javascript/api/excel/excel.documentpropertiesloadoptions#$all)||
||[Редактирование](/javascript/api/excel/excel.documentpropertiesloadoptions#author)|Получает или задает автора книги.|
||[категории](/javascript/api/excel/excel.documentpropertiesloadoptions#category)|Получает или задает категорию книги.|
||[comments](/javascript/api/excel/excel.documentpropertiesloadoptions#comments)|Получает или задает примечания книги.|
||[company](/javascript/api/excel/excel.documentpropertiesloadoptions#company)|Получает или задает компанию книги.|
||[creationDate](/javascript/api/excel/excel.documentpropertiesloadoptions#creationdate)|Получает дату создания книги. Только для чтения.|
||[keyword](/javascript/api/excel/excel.documentpropertiesloadoptions#keywords)|Получает или задает ключевые слова книги.|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesloadoptions#lastauthor)|Получает последнего автора книги. Только для чтения.|
||[manager](/javascript/api/excel/excel.documentpropertiesloadoptions#manager)|Получает или задает менеджера книги.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesloadoptions#revisionnumber)|Получает номер редакции книги. Только для чтения.|
||[subject](/javascript/api/excel/excel.documentpropertiesloadoptions#subject)|Получает или задает тему книги.|
||[заголовок](/javascript/api/excel/excel.documentpropertiesloadoptions#title)|Получает или задает название книги.|
|[Документпропертиесупдатедата](/javascript/api/excel/excel.documentpropertiesupdatedata)|[Редактирование](/javascript/api/excel/excel.documentpropertiesupdatedata#author)|Получает или задает автора книги.|
||[категории](/javascript/api/excel/excel.documentpropertiesupdatedata#category)|Получает или задает категорию книги.|
||[comments](/javascript/api/excel/excel.documentpropertiesupdatedata#comments)|Получает или задает примечания книги.|
||[company](/javascript/api/excel/excel.documentpropertiesupdatedata#company)|Получает или задает компанию книги.|
||[keyword](/javascript/api/excel/excel.documentpropertiesupdatedata#keywords)|Получает или задает ключевые слова книги.|
||[manager](/javascript/api/excel/excel.documentpropertiesupdatedata#manager)|Получает или задает менеджера книги.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesupdatedata#revisionnumber)|Получает номер редакции книги. Только для чтения.|
||[subject](/javascript/api/excel/excel.documentpropertiesupdatedata#subject)|Получает или задает тему книги.|
||[заголовок](/javascript/api/excel/excel.documentpropertiesupdatedata#title)|Получает или задает название книги.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Получает или задает формулу именованного элемента.  Формула всегда начинается со знака "=".|
||[Аррайвалуес](/javascript/api/excel/excel.nameditem#arrayvalues)|Возвращает объект, содержащий значения и типы именованного элемента. Только для чтения.|
|[Намедитемаррайвалуес](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Представляет типы для каждого элемента в именованном массиве элементов|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Представляет значения каждого элемента в массиве именованных элементов.|
|[Намедитемаррайвалуесдата](/javascript/api/excel/excel.nameditemarrayvaluesdata)|[types](/javascript/api/excel/excel.nameditemarrayvaluesdata#types)|Представляет типы для каждого элемента в именованном массиве элементов|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesdata#values)|Представляет значения каждого элемента в массиве именованных элементов.|
|[Намедитемаррайвалуеслоадоптионс](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions)|[$all](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#$all)||
||[types](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#types)|Представляет типы для каждого элемента в именованном массиве элементов|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#values)|Представляет значения каждого элемента в массиве именованных элементов.|
|[Намедитемколлектионлоадоптионс](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[Аррайвалуес](/javascript/api/excel/excel.nameditemcollectionloadoptions#arrayvalues)|Для каждого элемента в коллекции: Возвращает объект, содержащий значения и типы именованного элемента.|
||[formula](/javascript/api/excel/excel.nameditemcollectionloadoptions#formula)|Для каждого элемента в коллекции: Получает или задает формулу именованного элемента.  Формула всегда начинается со знака "=".|
|[Намедитемдата](/javascript/api/excel/excel.nameditemdata)|[Аррайвалуес](/javascript/api/excel/excel.nameditemdata#arrayvalues)|Возвращает объект, содержащий значения и типы именованного элемента. Только для чтения.|
||[formula](/javascript/api/excel/excel.nameditemdata#formula)|Получает или задает формулу именованного элемента.  Формула всегда начинается со знака "=".|
|[Намедитемлоадоптионс](/javascript/api/excel/excel.nameditemloadoptions)|[Аррайвалуес](/javascript/api/excel/excel.nameditemloadoptions#arrayvalues)|Возвращает объект, содержащий значения и типы именованного элемента.|
||[formula](/javascript/api/excel/excel.nameditemloadoptions#formula)|Получает или задает формулу именованного элемента.  Формула всегда начинается со знака "=".|
|[Намедитемупдатедата](/javascript/api/excel/excel.nameditemupdatedata)|[formula](/javascript/api/excel/excel.nameditemupdatedata#formula)|Получает или задает формулу именованного элемента.  Формула всегда начинается со знака "=".|
|[Range](/javascript/api/excel/excel.range)|[Жетабсолутересизедранже (Нумровс: число, Нумколумнс: число)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Получает объект Range с той же верхней левой ячейкой, что и текущий объект Range, но с указанным количеством строк и столбцов.|
||["-изображение" ()](/javascript/api/excel/excel.range#getimage--)|Отрисовывает диапазон в виде PNG-изображения в кодировке Base64.|
||[Жетсурраундингрегион ()](/javascript/api/excel/excel.range#getsurroundingregion--)|Возвращает объект Range, представляющий область вокруг верхней левой ячейки в этом диапазоне. Это диапазон, ограниченный любым сочетанием пустых строк и столбцов, относящихся к этому диапазону.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Представляет гиперссылку для текущего диапазона.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Представляет код числового формата Excel для указанного диапазона в виде строки на языке пользователя.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Указывает, является ли текущий диапазон целым столбцом. Только для чтения.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Указывает, является ли текущий диапазон целой строкой. Только для чтения.|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|Отображает карточку для активной ячейки, если она имеет содержимое c форматированным значением.|
||[style](/javascript/api/excel/excel.range#style)|Представляет стиль текущего диапазона.|
|[Ранжедата](/javascript/api/excel/excel.rangedata)|[hyperlink](/javascript/api/excel/excel.rangedata#hyperlink)|Представляет гиперссылку для текущего диапазона.|
||[isEntireColumn](/javascript/api/excel/excel.rangedata#isentirecolumn)|Указывает, является ли текущий диапазон целым столбцом. Только для чтения.|
||[isEntireRow](/javascript/api/excel/excel.rangedata#isentirerow)|Указывает, является ли текущий диапазон целой строкой. Только для чтения.|
||[numberFormatLocal](/javascript/api/excel/excel.rangedata#numberformatlocal)|Представляет код числового формата Excel для указанного диапазона в виде строки на языке пользователя.|
||[style](/javascript/api/excel/excel.rangedata#style)|Представляет стиль текущего диапазона.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Получает или задает ориентацию текста всех ячеек в диапазоне.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Определяет, равна ли высота строки объекта Range стандартной высоте листа.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Указывает, равняется ли ширина столбца объекта Range стандартной шириной листа.|
|[Ранжеформатдата](/javascript/api/excel/excel.rangeformatdata)|[textOrientation](/javascript/api/excel/excel.rangeformatdata#textorientation)|Получает или задает ориентацию текста всех ячеек в диапазоне.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatdata#usestandardheight)|Определяет, равна ли высота строки объекта Range стандартной высоте листа.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatdata#usestandardwidth)|Указывает, равняется ли ширина столбца объекта Range стандартной шириной листа.|
|[Ранжеформатлоадоптионс](/javascript/api/excel/excel.rangeformatloadoptions)|[textOrientation](/javascript/api/excel/excel.rangeformatloadoptions#textorientation)|Получает или задает ориентацию текста всех ячеек в диапазоне.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatloadoptions#usestandardheight)|Определяет, равна ли высота строки объекта Range стандартной высоте листа.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatloadoptions#usestandardwidth)|Указывает, равняется ли ширина столбца объекта Range стандартной шириной листа.|
|[Ранжеформатупдатедата](/javascript/api/excel/excel.rangeformatupdatedata)|[textOrientation](/javascript/api/excel/excel.rangeformatupdatedata#textorientation)|Получает или задает ориентацию текста всех ячеек в диапазоне.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatupdatedata#usestandardheight)|Определяет, равна ли высота строки объекта Range стандартной высоте листа.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatupdatedata#usestandardwidth)|Указывает, равняется ли ширина столбца объекта Range стандартной шириной листа.|
|[Ранжехиперлинк](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Представляет целевой URL-адрес для гиперссылки.|
||[Документреференце](/javascript/api/excel/excel.rangehyperlink#documentreference)|Представляет целевую ссылку на документ для гиперссылки.|
||[Сказок](/javascript/api/excel/excel.rangehyperlink#screentip)|Представляет строку, отображаемую при наведении указателя на гиперссылку.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Представляет строку, отображаемую в верхней левой ячейке диапазона.|
|[Ранжелоадоптионс](/javascript/api/excel/excel.rangeloadoptions)|[hyperlink](/javascript/api/excel/excel.rangeloadoptions#hyperlink)|Представляет гиперссылку для текущего диапазона.|
||[isEntireColumn](/javascript/api/excel/excel.rangeloadoptions#isentirecolumn)|Указывает, является ли текущий диапазон целым столбцом. Только для чтения.|
||[isEntireRow](/javascript/api/excel/excel.rangeloadoptions#isentirerow)|Указывает, является ли текущий диапазон целой строкой. Только для чтения.|
||[numberFormatLocal](/javascript/api/excel/excel.rangeloadoptions#numberformatlocal)|Представляет код числового формата Excel для указанного диапазона в виде строки на языке пользователя.|
||[style](/javascript/api/excel/excel.rangeloadoptions#style)|Представляет стиль текущего диапазона.|
|[Ранжеупдатедата](/javascript/api/excel/excel.rangeupdatedata)|[hyperlink](/javascript/api/excel/excel.rangeupdatedata#hyperlink)|Представляет гиперссылку для текущего диапазона.|
||[numberFormatLocal](/javascript/api/excel/excel.rangeupdatedata#numberformatlocal)|Представляет код числового формата Excel для указанного диапазона в виде строки на языке пользователя.|
||[style](/javascript/api/excel/excel.rangeupdatedata#style)|Представляет стиль текущего диапазона.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Удаляет этот стиль.|
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
||[Set (Properties: Excel. Style)](/javascript/api/excel/excel.style#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Стилеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.style#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Представляет вертикальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Указывает, применяет ли Microsoft Excel обтекание текстом для объекта.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Добавляет новый стиль в коллекцию.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Получает стиль по имени.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Стилеколлектионлоадоптионс](/javascript/api/excel/excel.stylecollectionloadoptions)|[$all](/javascript/api/excel/excel.stylecollectionloadoptions#$all)||
||[borders](/javascript/api/excel/excel.stylecollectionloadoptions#borders)|Для каждого элемента в коллекции: граница коллекции из четырех объектов Border, представляющих стиль четырех границ.|
||[builtIn](/javascript/api/excel/excel.stylecollectionloadoptions#builtin)|Для каждого элемента в коллекции: указывает, является ли этот стиль встроенным.|
||[fill](/javascript/api/excel/excel.stylecollectionloadoptions#fill)|Для каждого элемента в коллекции: заливка стиля.|
||[font](/javascript/api/excel/excel.stylecollectionloadoptions#font)|Для каждого элемента в коллекции: объект Font, представляющий шрифт стиля.|
||[formulaHidden](/javascript/api/excel/excel.stylecollectionloadoptions#formulahidden)|Для каждого элемента в коллекции: указывает, будет ли скрыта формула при защите листа.|
||[horizontalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#horizontalalignment)|Для каждого элемента в коллекции: представляет горизонтальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[includeAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#includealignment)|Для каждого элемента в коллекции: указывает, включают ли стили свойства Indent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel и TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.stylecollectionloadoptions#includeborder)|Для каждого элемента в коллекции: указывает, включают ли стили свойства границы цвета, ColorIndex, LineStyle и Weight.|
||[includeFont](/javascript/api/excel/excel.stylecollectionloadoptions#includefont)|Для каждого элемента в коллекции: указывает, включает ли стиль фон, полужирный шрифт, цвет, ColorIndex, FontStyle, курсив, имя, размер, зачеркивание, подстрочный знак, верхний индекс и подчеркивание шрифта.|
||[includeNumber](/javascript/api/excel/excel.stylecollectionloadoptions#includenumber)|Для каждого элемента в коллекции: указывает, содержит ли стиль свойство NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.stylecollectionloadoptions#includepatterns)|Для каждого элемента в коллекции: указывает, включают ли стили свойства Color, ColorIndex, InvertIfNegative, pattern, PatternColor и PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.stylecollectionloadoptions#includeprotection)|Для каждого элемента в коллекции: указывает, содержит ли стиль свойства FormulaHidden и locked Protection.|
||[indentLevel](/javascript/api/excel/excel.stylecollectionloadoptions#indentlevel)|Для каждого элемента в коллекции: целое число от 0 до 250, обозначающее уровень отступа для стиля.|
||[locked](/javascript/api/excel/excel.stylecollectionloadoptions#locked)|Для каждого элемента в коллекции: указывает, заблокирован ли объект, когда лист защищен.|
||[name](/javascript/api/excel/excel.stylecollectionloadoptions#name)|Для каждого элемента в коллекции: имя стиля.|
||[numberFormat](/javascript/api/excel/excel.stylecollectionloadoptions#numberformat)|Для каждого элемента в коллекции: код формата числового формата для стиля.|
||[numberFormatLocal](/javascript/api/excel/excel.stylecollectionloadoptions#numberformatlocal)|Для каждого элемента в коллекции: локализованный код формата числового формата для стиля.|
||[readingOrder](/javascript/api/excel/excel.stylecollectionloadoptions#readingorder)|Для каждого элемента в коллекции: порядок чтения для стиля.|
||[shrinkToFit](/javascript/api/excel/excel.stylecollectionloadoptions#shrinktofit)|Для каждого элемента в коллекции: указывает, сжимается ли текст автоматически в соответствии с шириной доступной ширины столбца.|
||[verticalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#verticalalignment)|Для каждого элемента в коллекции: представляет вертикальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.stylecollectionloadoptions#wraptext)|Для каждого элемента в коллекции: указывает, переносит ли Microsoft Excel текст в объекте.|
|[Стиледата](/javascript/api/excel/excel.styledata)|[borders](/javascript/api/excel/excel.styledata#borders)|Коллекция Border из четырех объектов Border, представляющих стиль четырех границ.|
||[builtIn](/javascript/api/excel/excel.styledata#builtin)|Указывает, является ли стиль встроенным.|
||[fill](/javascript/api/excel/excel.styledata#fill)|Заливка стиля.|
||[font](/javascript/api/excel/excel.styledata#font)|Объект Font, представляющий шрифт стиля.|
||[formulaHidden](/javascript/api/excel/excel.styledata#formulahidden)|Указывает, будет ли формула скрыта, если лист защищен.|
||[horizontalAlignment](/javascript/api/excel/excel.styledata#horizontalalignment)|Представляет горизонтальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[includeAlignment](/javascript/api/excel/excel.styledata#includealignment)|Указывает, содержатся ли в стиле такие свойства, как AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel и TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.styledata#includeborder)|Указывает, содержатся ли в стиле такие свойства границ, как Color, ColorIndex, LineStyle и Weight.|
||[includeFont](/javascript/api/excel/excel.styledata#includefont)|Указывает, содержатся ли в стиле такие свойства шрифта, как Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript и Underline.|
||[includeNumber](/javascript/api/excel/excel.styledata#includenumber)|Указывает, содержится ли в стиле свойство NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.styledata#includepatterns)|Указывает, содержатся ли в стиле такие внутренние свойства, как Color, ColorIndex, InvertIfNegative, Pattern, PatternColor и PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.styledata#includeprotection)|Указывает, содержатся ли в стиле такие свойства защиты, как FormulaHidden и Locked.|
||[indentLevel](/javascript/api/excel/excel.styledata#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа для стиля.|
||[locked](/javascript/api/excel/excel.styledata#locked)|Указывает, заблокирован ли объект, если лист защищен.|
||[name](/javascript/api/excel/excel.styledata#name)|Имя стиля.|
||[numberFormat](/javascript/api/excel/excel.styledata#numberformat)|Код числового формата для стиля.|
||[numberFormatLocal](/javascript/api/excel/excel.styledata#numberformatlocal)|Локализованный код числового формата для стиля.|
||[readingOrder](/javascript/api/excel/excel.styledata#readingorder)|Направление чтения для стиля.|
||[shrinkToFit](/javascript/api/excel/excel.styledata#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
||[verticalAlignment](/javascript/api/excel/excel.styledata#verticalalignment)|Представляет вертикальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.styledata#wraptext)|Указывает, применяет ли Microsoft Excel обтекание текстом для объекта.|
|[Стилелоадоптионс](/javascript/api/excel/excel.styleloadoptions)|[$all](/javascript/api/excel/excel.styleloadoptions#$all)||
||[borders](/javascript/api/excel/excel.styleloadoptions#borders)|Коллекция Border из четырех объектов Border, представляющих стиль четырех границ.|
||[builtIn](/javascript/api/excel/excel.styleloadoptions#builtin)|Указывает, является ли стиль встроенным.|
||[fill](/javascript/api/excel/excel.styleloadoptions#fill)|Заливка стиля.|
||[font](/javascript/api/excel/excel.styleloadoptions#font)|Объект Font, представляющий шрифт стиля.|
||[formulaHidden](/javascript/api/excel/excel.styleloadoptions#formulahidden)|Указывает, будет ли формула скрыта, если лист защищен.|
||[horizontalAlignment](/javascript/api/excel/excel.styleloadoptions#horizontalalignment)|Представляет горизонтальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[includeAlignment](/javascript/api/excel/excel.styleloadoptions#includealignment)|Указывает, содержатся ли в стиле такие свойства, как AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel и TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.styleloadoptions#includeborder)|Указывает, содержатся ли в стиле такие свойства границ, как Color, ColorIndex, LineStyle и Weight.|
||[includeFont](/javascript/api/excel/excel.styleloadoptions#includefont)|Указывает, содержатся ли в стиле такие свойства шрифта, как Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript и Underline.|
||[includeNumber](/javascript/api/excel/excel.styleloadoptions#includenumber)|Указывает, содержится ли в стиле свойство NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.styleloadoptions#includepatterns)|Указывает, содержатся ли в стиле такие внутренние свойства, как Color, ColorIndex, InvertIfNegative, Pattern, PatternColor и PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.styleloadoptions#includeprotection)|Указывает, содержатся ли в стиле такие свойства защиты, как FormulaHidden и Locked.|
||[indentLevel](/javascript/api/excel/excel.styleloadoptions#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа для стиля.|
||[locked](/javascript/api/excel/excel.styleloadoptions#locked)|Указывает, заблокирован ли объект, если лист защищен.|
||[name](/javascript/api/excel/excel.styleloadoptions#name)|Имя стиля.|
||[numberFormat](/javascript/api/excel/excel.styleloadoptions#numberformat)|Код числового формата для стиля.|
||[numberFormatLocal](/javascript/api/excel/excel.styleloadoptions#numberformatlocal)|Локализованный код числового формата для стиля.|
||[readingOrder](/javascript/api/excel/excel.styleloadoptions#readingorder)|Направление чтения для стиля.|
||[shrinkToFit](/javascript/api/excel/excel.styleloadoptions#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
||[verticalAlignment](/javascript/api/excel/excel.styleloadoptions#verticalalignment)|Представляет вертикальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.styleloadoptions#wraptext)|Указывает, применяет ли Microsoft Excel обтекание текстом для объекта.|
|[Стилеупдатедата](/javascript/api/excel/excel.styleupdatedata)|[borders](/javascript/api/excel/excel.styleupdatedata#borders)|Коллекция Border из четырех объектов Border, представляющих стиль четырех границ.|
||[fill](/javascript/api/excel/excel.styleupdatedata#fill)|Заливка стиля.|
||[font](/javascript/api/excel/excel.styleupdatedata#font)|Объект Font, представляющий шрифт стиля.|
||[formulaHidden](/javascript/api/excel/excel.styleupdatedata#formulahidden)|Указывает, будет ли формула скрыта, если лист защищен.|
||[horizontalAlignment](/javascript/api/excel/excel.styleupdatedata#horizontalalignment)|Представляет горизонтальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[includeAlignment](/javascript/api/excel/excel.styleupdatedata#includealignment)|Указывает, содержатся ли в стиле такие свойства, как AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel и TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.styleupdatedata#includeborder)|Указывает, содержатся ли в стиле такие свойства границ, как Color, ColorIndex, LineStyle и Weight.|
||[includeFont](/javascript/api/excel/excel.styleupdatedata#includefont)|Указывает, содержатся ли в стиле такие свойства шрифта, как Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript и Underline.|
||[includeNumber](/javascript/api/excel/excel.styleupdatedata#includenumber)|Указывает, содержится ли в стиле свойство NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.styleupdatedata#includepatterns)|Указывает, содержатся ли в стиле такие внутренние свойства, как Color, ColorIndex, InvertIfNegative, Pattern, PatternColor и PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.styleupdatedata#includeprotection)|Указывает, содержатся ли в стиле такие свойства защиты, как FormulaHidden и Locked.|
||[indentLevel](/javascript/api/excel/excel.styleupdatedata#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа для стиля.|
||[locked](/javascript/api/excel/excel.styleupdatedata#locked)|Указывает, заблокирован ли объект, если лист защищен.|
||[numberFormat](/javascript/api/excel/excel.styleupdatedata#numberformat)|Код числового формата для стиля.|
||[numberFormatLocal](/javascript/api/excel/excel.styleupdatedata#numberformatlocal)|Локализованный код числового формата для стиля.|
||[readingOrder](/javascript/api/excel/excel.styleupdatedata#readingorder)|Направление чтения для стиля.|
||[shrinkToFit](/javascript/api/excel/excel.styleupdatedata#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
||[verticalAlignment](/javascript/api/excel/excel.styleupdatedata#verticalalignment)|Представляет вертикальное выравнивание для стиля. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.styleupdatedata#wraptext)|Указывает, применяет ли Microsoft Excel обтекание текстом для объекта.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Происходит при изменении данных в ячейках в определенной таблице.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Возникает при изменении выбора в определенной таблице.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Получает адрес, представляющий измененную область таблицы на конкретном листе.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Получает тип изменения, представляющий способ запуска события Changed. Дополнительные сведения см. в статье Excel. Датачанжетипе.|
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
|[Воркбукдата](/javascript/api/excel/excel.workbookdata)|[name](/javascript/api/excel/excel.workbookdata#name)|Получает имя книги. Только для чтения.|
||[properties](/javascript/api/excel/excel.workbookdata#properties)|Получает свойства книги. Только для чтения.|
||[protection](/javascript/api/excel/excel.workbookdata#protection)|Возвращает объект защиты книги. Только для чтения.|
||[стили](/javascript/api/excel/excel.workbookdata#styles)|Представляет коллекцию стилей, связанных с книгой. Только для чтения.|
|[Воркбуклоадоптионс](/javascript/api/excel/excel.workbookloadoptions)|[name](/javascript/api/excel/excel.workbookloadoptions#name)|Получает имя книги. Только для чтения.|
||[properties](/javascript/api/excel/excel.workbookloadoptions#properties)|Получает свойства книги.|
||[protection](/javascript/api/excel/excel.workbookloadoptions#protection)|Возвращает объект защиты книги.|
|[Воркбукпротектион](/javascript/api/excel/excel.workbookprotection)|[Защита (пароль?: строка)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Защищает книгу. Выдает ошибку, если книга защищена.|
||[Защита](/javascript/api/excel/excel.workbookprotection#protected)|Указывает, защищена ли книга. Только для чтения.|
||[снять защиту (пароль?: строка)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Снимает защиту с книги.|
|[Воркбукпротектиондата](/javascript/api/excel/excel.workbookprotectiondata)|[Защита](/javascript/api/excel/excel.workbookprotectiondata#protected)|Указывает, защищена ли книга. Только для чтения.|
|[Воркбукпротектионлоадоптионс](/javascript/api/excel/excel.workbookprotectionloadoptions)|[$all](/javascript/api/excel/excel.workbookprotectionloadoptions#$all)||
||[Защита](/javascript/api/excel/excel.workbookprotectionloadoptions#protected)|Указывает, защищена ли книга. Только для чтения.|
|[Воркбукупдатедата](/javascript/api/excel/excel.workbookupdatedata)|[properties](/javascript/api/excel/excel.workbookupdatedata#properties)|Получает свойства книги.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Copy (Поситионтипе?: "None" \| "Before" \| "после" \| "начало \| ", релативето?: Excel. лист)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Копирует лист и размещает его в указанном положении. Возвращает скопированный лист.|
||[Copy (Поситионтипе?: Excel. Воркшитпоситионтипе, Релативето?: Excel. лист)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Копирует лист и размещает его в указанном положении. Возвращает скопированный лист.|
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
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Возникает при активации любого листа в книге.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Возникает при добавлении нового листа в книгу.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Возникает, когда отключается любой лист в книге.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Возникает при удалении листа из книги.|
|[Воркшитколлектионлоадоптионс](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardheight)|Для каждого элемента в коллекции: возвращается Стандартная высота (по умолчанию) всех строк на листе в пунктах. Только для чтения.|
||[standardWidth](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardwidth)|Для каждого элемента в коллекции: Возвращает или задает стандартную ширину (по умолчанию) всех столбцов на листе.|
||[Табколор](/javascript/api/excel/excel.worksheetcollectionloadoptions#tabcolor)|Для каждого элемента в коллекции: Получает или задает цвет ярлычка листа.|
|[Воркшитдата](/javascript/api/excel/excel.worksheetdata)|[standardHeight](/javascript/api/excel/excel.worksheetdata#standardheight)|Возвращает стандартную (по умолчанию) высоту всех строк на листе (в пунктах). Только для чтения.|
||[standardWidth](/javascript/api/excel/excel.worksheetdata#standardwidth)|Возвращает или задает стандартную (по умолчанию) ширину всех столбцов на листе.|
||[Табколор](/javascript/api/excel/excel.worksheetdata#tabcolor)|Получает или задает цвет вкладки листа.|
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
|[Воркшитлоадоптионс](/javascript/api/excel/excel.worksheetloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetloadoptions#standardheight)|Возвращает стандартную (по умолчанию) высоту всех строк на листе (в пунктах). Только для чтения.|
||[standardWidth](/javascript/api/excel/excel.worksheetloadoptions#standardwidth)|Возвращает или задает стандартную (по умолчанию) ширину всех столбцов на листе.|
||[Табколор](/javascript/api/excel/excel.worksheetloadoptions#tabcolor)|Получает или задает цвет вкладки листа.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[снять защиту (пароль?: строка)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Снимает защиту с листа.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[Алловедитобжектс](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Представляет параметр защиты листа, разрешающий редактирование объектов.|
||[Алловедитсценариос](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Представляет параметр защиты листа, разрешающий редактирование сценариев.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Представляет параметр защиты рабочего листа для режима выделения.|
|[Воркшитселектиончанжедевентаргс](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Получает адрес диапазона, представляющий выделенную область конкретного листа.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменено выделение.|
|[Воркшитупдатедата](/javascript/api/excel/excel.worksheetupdatedata)|[standardWidth](/javascript/api/excel/excel.worksheetupdatedata#standardwidth)|Возвращает или задает стандартную (по умолчанию) ширину всех столбцов на листе.|
||[Табколор](/javascript/api/excel/excel.worksheetupdatedata#tabcolor)|Получает или задает цвет вкладки листа.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
