---
title: Набор обязательных элементов API JavaScript для Excel 1,1
description: Сведения о наборе требований ExcelApi 1,1
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 921a67b4242150d767fdac057d21c6fc510d98b3
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772053"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Набор обязательных элементов API JavaScript для Excel 1,1

API JavaScript для Excel 1,1 — это первая версия API. Это единственный набор обязательных элементов Excel, поддерживаемый Excel 2016.

## <a name="api-list"></a>Список API

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[Calculate (Калкулатионтипе: "Recalculate" \| Full " \| " фуллребуилд ")](/javascript/api/excel/excel.application#calculate-calculationtype-)|Пересчитывает данные во всех открытых в текущий момент книгах Excel.|
||[Calculate (Калкулатионтипе: Excel. Калкулатионтипе)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Пересчитывает данные во всех открытых в текущий момент книгах Excel.|
||[Калкулатионмоде](/javascript/api/excel/excel.application#calculationmode)|Возвращает режим вычислений, используемый в книге в соответствии с константами в Excel. Калкулатионмоде. Возможные значения: `Automatic`, где Excel управляет пересчетом; `AutomaticExceptTables`, где Excel контролирует пересчет, но игнорирует изменения в таблицах; `Manual`, где выполняется расчет, когда пользователь запрашивает его.|
||[Set (Properties: Excel. Application)](/javascript/api/excel/excel.application#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Аппликатионупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.application#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Аппликатиондата](/javascript/api/excel/excel.applicationdata)|[Калкулатионмоде](/javascript/api/excel/excel.applicationdata#calculationmode)|Возвращает режим вычислений, используемый в книге в соответствии с константами в Excel. Калкулатионмоде. Возможные значения: `Automatic`, где Excel управляет пересчетом; `AutomaticExceptTables`, где Excel контролирует пересчет, но игнорирует изменения в таблицах; `Manual`, где выполняется расчет, когда пользователь запрашивает его.|
|[Аппликатионлоадоптионс](/javascript/api/excel/excel.applicationloadoptions)|[$all](/javascript/api/excel/excel.applicationloadoptions#$all)||
||[Калкулатионмоде](/javascript/api/excel/excel.applicationloadoptions#calculationmode)|Возвращает режим вычислений, используемый в книге в соответствии с константами в Excel. Калкулатионмоде. Возможные значения: `Automatic`, где Excel управляет пересчетом; `AutomaticExceptTables`, где Excel контролирует пересчет, но игнорирует изменения в таблицах; `Manual`, где выполняется расчет, когда пользователь запрашивает его.|
|[Аппликатионупдатедата](/javascript/api/excel/excel.applicationupdatedata)|[Калкулатионмоде](/javascript/api/excel/excel.applicationupdatedata#calculationmode)|Возвращает режим вычислений, используемый в книге в соответствии с константами в Excel. Калкулатионмоде. Возможные значения: `Automatic`, где Excel управляет пересчетом; `AutomaticExceptTables`, где Excel контролирует пересчет, но игнорирует изменения в таблицах; `Manual`, где выполняется расчет, когда пользователь запрашивает его.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Возвращает представленный привязкой диапазон. Если тип привязки неправильный, выдается ошибка.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Возвращает представленную привязкой таблицу. Если тип привязки неправильный, выдается ошибка.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Возвращает представленный привязкой текст. Если тип привязки неправильный, выдается ошибка.|
||[id](/javascript/api/excel/excel.binding#id)|Представляет идентификатор привязки. Только для чтения.|
||[type](/javascript/api/excel/excel.binding#type)|Возвращает тип привязки. Дополнительные сведения см. в статье Excel. BindingType. Только для чтения.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Возвращает объект привязки по идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Возвращает объект привязки с учетом его положения в массиве элементов.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Возвращает число привязок в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Биндингколлектионлоадоптионс](/javascript/api/excel/excel.bindingcollectionloadoptions)|[$all](/javascript/api/excel/excel.bindingcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingcollectionloadoptions#id)|Для каждого элемента в коллекции: представляет идентификатор привязки. Только для чтения.|
||[type](/javascript/api/excel/excel.bindingcollectionloadoptions#type)|Для каждого элемента в коллекции: Возвращает тип привязки. Дополнительные сведения см. в статье Excel. BindingType. Только для чтения.|
|[Биндингдата](/javascript/api/excel/excel.bindingdata)|[id](/javascript/api/excel/excel.bindingdata#id)|Представляет идентификатор привязки. Только для чтения.|
||[type](/javascript/api/excel/excel.bindingdata#type)|Возвращает тип привязки. Дополнительные сведения см. в статье Excel. BindingType. Только для чтения.|
|[Биндинглоадоптионс](/javascript/api/excel/excel.bindingloadoptions)|[$all](/javascript/api/excel/excel.bindingloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingloadoptions#id)|Представляет идентификатор привязки. Только для чтения.|
||[type](/javascript/api/excel/excel.bindingloadoptions#type)|Возвращает тип привязки. Дополнительные сведения см. в статье Excel. BindingType. Только для чтения.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Удаляет объект диаграммы.|
||[height](/javascript/api/excel/excel.chart#height)|Обозначает высоту объекта диаграммы (в пунктах).|
||[left](/javascript/api/excel/excel.chart#left)|Расстояние в пунктах от левого края диаграммы до начала листа.|
||[name](/javascript/api/excel/excel.chart#name)|Обозначает имя объекта диаграммы.|
||[Axes](/javascript/api/excel/excel.chart#axes)|Представляет оси диаграммы. Только для чтения.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Представляет метки данных на диаграмме. Только для чтения.|
||[format](/javascript/api/excel/excel.chart#format)|Инкапсулирует свойства формата для области диаграммы. Только для чтения.|
||[списком](/javascript/api/excel/excel.chart#legend)|Представляет условные обозначения для диаграммы. Только для чтения.|
||[series](/javascript/api/excel/excel.chart#series)|Представляет один ряд данных или коллекцию рядов данных в диаграмме. Только для чтения.|
||[заголовок](/javascript/api/excel/excel.chart#title)|Представляет заголовок указанной диаграммы, включая его текст, видимость, положение и форматирование. Только для чтения.|
||[Set (Properties: Excel. Chart)](/javascript/api/excel/excel.chart#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chart#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[setData (sourceData: Range, seriesBy?: "Auto" \| "Columns \| " "Rows")](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Сбрасывает исходные данные для диаграммы.|
||[setData (sourceData: Range, seriesBy?: Excel. Чартсериесби)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Сбрасывает исходные данные для диаграммы.|
||[setPosition (startCell: строка \| диапазона, endCell?: строка \| диапазона)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Располагает диаграмму относительно ячеек на листе.|
||[top](/javascript/api/excel/excel.chart#top)|Представляет расстояние в пунктах от верхнего края объекта до верхнего края первой строки (на листе) или до верхнего края области диаграммы (на диаграмме).|
||[width](/javascript/api/excel/excel.chart#width)|Представляет ширину объекта диаграммы (в пунктах).|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона. Только для чтения.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для текущего объекта. Только для чтения.|
||[Set (Properties: Excel. ChartAreaFormat)](/javascript/api/excel/excel.chartareaformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартареаформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartareaformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартареаформатдата](/javascript/api/excel/excel.chartareaformatdata)|[font](/javascript/api/excel/excel.chartareaformatdata#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для текущего объекта. Только для чтения.|
|[Чартареаформатлоадоптионс](/javascript/api/excel/excel.chartareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartareaformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartareaformatloadoptions#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для текущего объекта.|
|[Чартареаформатупдатедата](/javascript/api/excel/excel.chartareaformatupdatedata)|[font](/javascript/api/excel/excel.chartareaformatupdatedata#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для текущего объекта.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[Категоряксис](/javascript/api/excel/excel.chartaxes#categoryaxis)|Представляет ось категорий на диаграмме. Только для чтения.|
||[Сериесаксис](/javascript/api/excel/excel.chartaxes#seriesaxis)|Представляет ось ряда данных для объемной диаграммы. Только для чтения.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Представляет ось значений для оси. Только для чтения.|
||[Set (Properties: Excel. ChartAxes)](/javascript/api/excel/excel.chartaxes#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартаксесупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxes#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартаксесдата](/javascript/api/excel/excel.chartaxesdata)|[Категоряксис](/javascript/api/excel/excel.chartaxesdata#categoryaxis)|Представляет ось категорий на диаграмме. Только для чтения.|
||[Сериесаксис](/javascript/api/excel/excel.chartaxesdata#seriesaxis)|Представляет ось ряда данных для объемной диаграммы. Только для чтения.|
||[valueAxis](/javascript/api/excel/excel.chartaxesdata#valueaxis)|Представляет ось значений для оси. Только для чтения.|
|[Чартаксеслоадоптионс](/javascript/api/excel/excel.chartaxesloadoptions)|[$all](/javascript/api/excel/excel.chartaxesloadoptions#$all)||
||[Категоряксис](/javascript/api/excel/excel.chartaxesloadoptions#categoryaxis)|Представляет ось категорий на диаграмме.|
||[Сериесаксис](/javascript/api/excel/excel.chartaxesloadoptions#seriesaxis)|Представляет ось ряда данных для объемной диаграммы.|
||[valueAxis](/javascript/api/excel/excel.chartaxesloadoptions#valueaxis)|Представляет ось значений для оси.|
|[Чартаксесупдатедата](/javascript/api/excel/excel.chartaxesupdatedata)|[Категоряксис](/javascript/api/excel/excel.chartaxesupdatedata#categoryaxis)|Представляет ось категорий на диаграмме.|
||[Сериесаксис](/javascript/api/excel/excel.chartaxesupdatedata#seriesaxis)|Представляет ось ряда данных для объемной диаграммы.|
||[valueAxis](/javascript/api/excel/excel.chartaxesupdatedata#valueaxis)|Представляет ось значений для оси.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Обозначает интервал между двумя основными делениями. Можно указать в виде числового значения или пустой строки.  Возвращаемое значение всегда является числом.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Представляет максимальное значение на оси значений.  Можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси).  Возвращаемое значение всегда является числом.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Представляет минимальное значение на оси значений. Ему можно присвоить числовое значение или пустую строку (для автоматически заданных значений оси). Всегда возвращает числовое значение.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Представляет интервал между двумя промежуточными делениями. Его можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси). Возвращаемое значение всегда является числом.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Представляет форматирование объекта диаграммы, в том числе форматирование линий и шрифта. Только для чтения.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Возвращает объект линии сетки, который представляет основные линии сетки для указанной оси. Только для чтения.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Возвращает объект Gridlines, который представляет вспомогательные линии сетки для указанной оси. Только для чтения.|
||[заголовок](/javascript/api/excel/excel.chartaxis#title)|Обозначает название оси. Только для чтения.|
||[Set (Properties: Excel. ChartAxis)](/javascript/api/excel/excel.chartaxis#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартаксисупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxis#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартаксисдата](/javascript/api/excel/excel.chartaxisdata)|[format](/javascript/api/excel/excel.chartaxisdata#format)|Представляет форматирование объекта диаграммы, в том числе форматирование линий и шрифта. Только для чтения.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisdata#majorgridlines)|Возвращает объект линии сетки, который представляет основные линии сетки для указанной оси. Только для чтения.|
||[majorUnit](/javascript/api/excel/excel.chartaxisdata#majorunit)|Обозначает интервал между двумя основными делениями. Можно указать в виде числового значения или пустой строки.  Возвращаемое значение всегда является числом.|
||[maximum](/javascript/api/excel/excel.chartaxisdata#maximum)|Представляет максимальное значение на оси значений.  Можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси).  Возвращаемое значение всегда является числом.|
||[minimum](/javascript/api/excel/excel.chartaxisdata#minimum)|Представляет минимальное значение на оси значений. Ему можно присвоить числовое значение или пустую строку (для автоматически заданных значений оси). Всегда возвращает числовое значение.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisdata#minorgridlines)|Возвращает объект Gridlines, который представляет вспомогательные линии сетки для указанной оси. Только для чтения.|
||[minorUnit](/javascript/api/excel/excel.chartaxisdata#minorunit)|Представляет интервал между двумя промежуточными делениями. Его можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси). Возвращаемое значение всегда является числом.|
||[заголовок](/javascript/api/excel/excel.chartaxisdata#title)|Обозначает название оси. Только для чтения.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для элемента оси диаграммы. Только для чтения.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Представляет форматирование линий диаграммы. Только для чтения.|
||[Set (Properties: Excel. ChartAxisFormat)](/javascript/api/excel/excel.chartaxisformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартаксисформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxisformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартаксисформатдата](/javascript/api/excel/excel.chartaxisformatdata)|[font](/javascript/api/excel/excel.chartaxisformatdata#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для элемента оси диаграммы. Только для чтения.|
||[line](/javascript/api/excel/excel.chartaxisformatdata#line)|Представляет форматирование линий диаграммы. Только для чтения.|
|[Чартаксисформатлоадоптионс](/javascript/api/excel/excel.chartaxisformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxisformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartaxisformatloadoptions#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для элемента оси диаграммы.|
||[line](/javascript/api/excel/excel.chartaxisformatloadoptions#line)|Представляет форматирование линий диаграммы.|
|[Чартаксисформатупдатедата](/javascript/api/excel/excel.chartaxisformatupdatedata)|[font](/javascript/api/excel/excel.chartaxisformatupdatedata#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для элемента оси диаграммы.|
||[line](/javascript/api/excel/excel.chartaxisformatupdatedata#line)|Представляет форматирование линий диаграммы.|
|[Чартаксислоадоптионс](/javascript/api/excel/excel.chartaxisloadoptions)|[$all](/javascript/api/excel/excel.chartaxisloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxisloadoptions#format)|Представляет форматирование объекта диаграммы, в том числе форматирование линий и шрифта.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#majorgridlines)|Возвращает объект линии сетки, который представляет основные линии сетки для указанной оси.|
||[majorUnit](/javascript/api/excel/excel.chartaxisloadoptions#majorunit)|Обозначает интервал между двумя основными делениями. Можно указать в виде числового значения или пустой строки.  Возвращаемое значение всегда является числом.|
||[maximum](/javascript/api/excel/excel.chartaxisloadoptions#maximum)|Представляет максимальное значение на оси значений.  Можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси).  Возвращаемое значение всегда является числом.|
||[minimum](/javascript/api/excel/excel.chartaxisloadoptions#minimum)|Представляет минимальное значение на оси значений. Ему можно присвоить числовое значение или пустую строку (для автоматически заданных значений оси). Всегда возвращает числовое значение.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#minorgridlines)|Возвращает объект Gridlines, который представляет вспомогательные линии сетки для указанной оси.|
||[minorUnit](/javascript/api/excel/excel.chartaxisloadoptions#minorunit)|Представляет интервал между двумя промежуточными делениями. Его можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси). Возвращаемое значение всегда является числом.|
||[заголовок](/javascript/api/excel/excel.chartaxisloadoptions#title)|Обозначает название оси.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Представляет форматирование для названия оси диаграммы. Только для чтения.|
||[Set (Properties: Excel. ChartAxisTitle)](/javascript/api/excel/excel.chartaxistitle#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартаксиститлеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxistitle#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Обозначает название оси.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Логическое значение, которое определяет видимость названия оси.|
|[Чартаксиститледата](/javascript/api/excel/excel.chartaxistitledata)|[format](/javascript/api/excel/excel.chartaxistitledata#format)|Представляет форматирование для названия оси диаграммы. Только для чтения.|
||[text](/javascript/api/excel/excel.chartaxistitledata#text)|Обозначает название оси.|
||[visible](/javascript/api/excel/excel.chartaxistitledata#visible)|Логическое значение, которое определяет видимость названия оси.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. объект заголовка оси диаграммы. Только для чтения.|
||[Set (Properties: Excel. ChartAxisTitleFormat)](/javascript/api/excel/excel.chartaxistitleformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартаксиститлеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxistitleformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартаксиститлеформатдата](/javascript/api/excel/excel.chartaxistitleformatdata)|[font](/javascript/api/excel/excel.chartaxistitleformatdata#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. объект заголовка оси диаграммы. Только для чтения.|
|[Чартаксиститлеформатлоадоптионс](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartaxistitleformatloadoptions#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. объект заголовка оси диаграммы.|
|[Чартаксиститлеформатупдатедата](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[font](/javascript/api/excel/excel.chartaxistitleformatupdatedata#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. объект заголовка оси диаграммы.|
|[Чартаксиститлелоадоптионс](/javascript/api/excel/excel.chartaxistitleloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxistitleloadoptions#format)|Представляет форматирование для названия оси диаграммы.|
||[text](/javascript/api/excel/excel.chartaxistitleloadoptions#text)|Обозначает название оси.|
||[visible](/javascript/api/excel/excel.chartaxistitleloadoptions#visible)|Логическое значение, которое определяет видимость названия оси.|
|[Чартаксиститлеупдатедата](/javascript/api/excel/excel.chartaxistitleupdatedata)|[format](/javascript/api/excel/excel.chartaxistitleupdatedata#format)|Представляет форматирование для названия оси диаграммы.|
||[text](/javascript/api/excel/excel.chartaxistitleupdatedata#text)|Обозначает название оси.|
||[visible](/javascript/api/excel/excel.chartaxistitleupdatedata#visible)|Логическое значение, которое определяет видимость названия оси.|
|[Чартаксисупдатедата](/javascript/api/excel/excel.chartaxisupdatedata)|[format](/javascript/api/excel/excel.chartaxisupdatedata#format)|Представляет форматирование объекта диаграммы, в том числе форматирование линий и шрифта.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#majorgridlines)|Возвращает объект линии сетки, который представляет основные линии сетки для указанной оси.|
||[majorUnit](/javascript/api/excel/excel.chartaxisupdatedata#majorunit)|Обозначает интервал между двумя основными делениями. Можно указать в виде числового значения или пустой строки.  Возвращаемое значение всегда является числом.|
||[maximum](/javascript/api/excel/excel.chartaxisupdatedata#maximum)|Представляет максимальное значение на оси значений.  Можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси).  Возвращаемое значение всегда является числом.|
||[minimum](/javascript/api/excel/excel.chartaxisupdatedata#minimum)|Представляет минимальное значение на оси значений. Ему можно присвоить числовое значение или пустую строку (для автоматически заданных значений оси). Всегда возвращает числовое значение.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#minorgridlines)|Возвращает объект Gridlines, который представляет вспомогательные линии сетки для указанной оси.|
||[minorUnit](/javascript/api/excel/excel.chartaxisupdatedata#minorunit)|Представляет интервал между двумя промежуточными делениями. Его можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси). Возвращаемое значение всегда является числом.|
||[заголовок](/javascript/api/excel/excel.chartaxisupdatedata#title)|Обозначает название оси.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[Add (Type \| : "Invalid" "ColumnClustered" \| "колумнстаккед" \| "ColumnStacked100" \| "3DColumnClustered" \| "3DColumnStacked" \| "3DColumnStacked100" \| "барклустеред" \| "барстаккед" \| "BarStacked100" \| "3DBarClustered" \| "3DBarStacked" \| "3DBarStacked100" \| "линестаккед" \| "LineStacked100" \| "линемаркерс" \| "линемаркерсстаккед" \| " LineMarkersStacked100 " \| " пиеофпие " \| " пииксплодед " \| " 3DPieExploded " \| " барофпие " \| " ксискаттерсмус " \| " ксискаттерсмусномаркерс " \| " ксискаттерлинес " \| " Ксискаттерлинесномаркерс " \| " ареастаккед " \| " AreaStacked100 " \| " 3DAreaStacked " \| " 3DAreaStacked100 " \| " кольцевых " \| " радармаркерс " \| " радарфиллед " \| " Поверхность " \| " сурфацевирефраме " \| " сурфацетопвиев " \| " сурфацетопвиеввирефраме " \| " пузырь " \| " Bubble3DEffect " \| " стоккхлк " \| " стоккохлк " \| " стокквхлк " \| " Стокквохлк " \| " цилиндерколклустеред " \| " цилиндерколстаккед " \| " CylinderColStacked100 " \| " цилиндербарклустеред " \| " цилиндербарстаккед " \| " CylinderBarStacked100 " \| " Цилиндеркол " \| " конеколклустеред " \| " конеколстаккед " \| " ConeColStacked100 " \| " конебарклустеред " \| " конебарстаккед " \| " ConeBarStacked100 " \| " конекол " \| " Пирамидколклустеред " \| " пирамидколстаккед " \| " PyramidColStacked100 " \| " пирамидбарклустеред " \| " пирамидбарстаккед " \| " PyramidBarStacked100 " \| " пирамидкол " \| " 3DColumn " \| "Line" \| "3DLine" \| "3DPie" \| "круг" \| "ксискаттер" \| "3DArea" \| "площадь" \| "кольцевой \| " "лепестк \| " "Гистограмма \| " " \| боксвхискер" " Парето " \| " регионмап " \| " Эта " \| " Каскад " \| " \| , "воронка", sourceData: Range, seriesBy?: "Auto" \| "Columns" \| "Rows")](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Создает диаграмму.|
||[Добавить (тип: Excel. ChartType, sourceData: Range, seriesBy?: Excel. Чартсериесби)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Создает диаграмму.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Возвращает диаграмму по ее имени. Если одно и то же имя принадлежит нескольким диаграммам, будет возвращена первая из них.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Возвращает диаграмму с учетом ее положения в коллекции.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Возвращает количество диаграмм на листе. Только для чтения.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Чартколлектионлоадоптионс](/javascript/api/excel/excel.chartcollectionloadoptions)|[$all](/javascript/api/excel/excel.chartcollectionloadoptions#$all)||
||[Axes](/javascript/api/excel/excel.chartcollectionloadoptions#axes)|Для каждого элемента в коллекции: представляет оси диаграммы.|
||[dataLabels](/javascript/api/excel/excel.chartcollectionloadoptions#datalabels)|Для каждого элемента в коллекции: представляет метки DataItem на диаграмме.|
||[format](/javascript/api/excel/excel.chartcollectionloadoptions#format)|Для каждого элемента в коллекции: инкапсулирует свойства формата для области диаграммы.|
||[height](/javascript/api/excel/excel.chartcollectionloadoptions#height)|Для каждого элемента в коллекции: представляет высоту объекта диаграммы (в пунктах).|
||[left](/javascript/api/excel/excel.chartcollectionloadoptions#left)|Для каждого элемента в коллекции: расстояние (в пунктах) от левого края диаграммы до начала листа.|
||[списком](/javascript/api/excel/excel.chartcollectionloadoptions#legend)|Для каждого элемента в коллекции: представляет условные обозначения для диаграммы.|
||[name](/javascript/api/excel/excel.chartcollectionloadoptions#name)|Для каждого элемента в коллекции: представляет имя объекта Chart.|
||[series](/javascript/api/excel/excel.chartcollectionloadoptions#series)|Для каждого элемента в коллекции: представляет один ряд или коллекцию рядов в диаграмме.|
||[заголовок](/javascript/api/excel/excel.chartcollectionloadoptions#title)|Для каждого элемента в коллекции: представляет название указанной диаграммы, включая текст, видимость, положение и форматирование заголовка.|
||[top](/javascript/api/excel/excel.chartcollectionloadoptions#top)|Для каждого элемента в коллекции: представляет расстояние (в пунктах) от верхнего края объекта до верхнего края строки 1 (на листе) или сверху области диаграммы (на диаграмме).|
||[width](/javascript/api/excel/excel.chartcollectionloadoptions#width)|Для каждого элемента в коллекции: представляет ширину (в пунктах) объекта Chart.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[Axes](/javascript/api/excel/excel.chartdata#axes)|Представляет оси диаграммы. Только для чтения.|
||[dataLabels](/javascript/api/excel/excel.chartdata#datalabels)|Представляет метки данных на диаграмме. Только для чтения.|
||[format](/javascript/api/excel/excel.chartdata#format)|Инкапсулирует свойства формата для области диаграммы. Только для чтения.|
||[height](/javascript/api/excel/excel.chartdata#height)|Обозначает высоту объекта диаграммы (в пунктах).|
||[left](/javascript/api/excel/excel.chartdata#left)|Расстояние в пунктах от левого края диаграммы до начала листа.|
||[списком](/javascript/api/excel/excel.chartdata#legend)|Представляет условные обозначения для диаграммы. Только для чтения.|
||[name](/javascript/api/excel/excel.chartdata#name)|Обозначает имя объекта диаграммы.|
||[series](/javascript/api/excel/excel.chartdata#series)|Представляет один ряд данных или коллекцию рядов данных в диаграмме. Только для чтения.|
||[заголовок](/javascript/api/excel/excel.chartdata#title)|Представляет заголовок указанной диаграммы, включая его текст, видимость, положение и форматирование. Только для чтения.|
||[top](/javascript/api/excel/excel.chartdata#top)|Представляет расстояние в пунктах от верхнего края объекта до верхнего края первой строки (на листе) или до верхнего края области диаграммы (на диаграмме).|
||[width](/javascript/api/excel/excel.chartdata#width)|Представляет ширину объекта диаграммы (в пунктах).|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Представляет формат заливки для текущей метки данных диаграммы. Только для чтения.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для подписи данных диаграммы. Только для чтения.|
||[Set (Properties: Excel. ChartDataLabelFormat)](/javascript/api/excel/excel.chartdatalabelformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартдаталабелформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartdatalabelformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартдаталабелформатдата](/javascript/api/excel/excel.chartdatalabelformatdata)|[font](/javascript/api/excel/excel.chartdatalabelformatdata#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для подписи данных диаграммы. Только для чтения.|
|[Чартдаталабелформатлоадоптионс](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartdatalabelformatloadoptions#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для подписи данных диаграммы.|
|[Чартдаталабелформатупдатедата](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[font](/javascript/api/excel/excel.chartdatalabelformatupdatedata#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для подписи данных диаграммы.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Представляет формат меток данных диаграммы, включая форматирование заливки и шрифтов. Только для чтения.|
||[символ](/javascript/api/excel/excel.chartdatalabels#separator)|Строка, представляющая разделитель, который используется для меток данных на диаграмме.|
||[Set (Properties: Excel. ChartDataLabels)](/javascript/api/excel/excel.chartdatalabels#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартдаталабелсупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartdatalabels#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[Чартдаталабелсдата](/javascript/api/excel/excel.chartdatalabelsdata)|[format](/javascript/api/excel/excel.chartdatalabelsdata#format)|Представляет формат меток данных диаграммы, включая форматирование заливки и шрифтов. Только для чтения.|
||[position](/javascript/api/excel/excel.chartdatalabelsdata#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[символ](/javascript/api/excel/excel.chartdatalabelsdata#separator)|Строка, представляющая разделитель, который используется для меток данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsdata#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsdata#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsdata#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsdata#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsdata#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsdata#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[Чартдаталабелслоадоптионс](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelsloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartdatalabelsloadoptions#format)|Представляет формат меток данных диаграммы, включая форматирование заливки и шрифтов.|
||[position](/javascript/api/excel/excel.chartdatalabelsloadoptions#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[символ](/javascript/api/excel/excel.chartdatalabelsloadoptions#separator)|Строка, представляющая разделитель, который используется для меток данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsloadoptions#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsloadoptions#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsloadoptions#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsloadoptions#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[Чартдаталабелсупдатедата](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[format](/javascript/api/excel/excel.chartdatalabelsupdatedata#format)|Представляет формат меток данных диаграммы, включая форматирование заливки и шрифтов.|
||[position](/javascript/api/excel/excel.chartdatalabelsupdatedata#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[символ](/javascript/api/excel/excel.chartdatalabelsupdatedata#separator)|Строка, представляющая разделитель, который используется для меток данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsupdatedata#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsupdatedata#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsupdatedata#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsupdatedata#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Очищает цвет заливки элемента диаграммы.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Устанавливает форматирование заливки элемента диаграммы на единый цвет.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.chartfont#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.chartfont#name)|Имя шрифта (например, Calibri)|
||[Set (Properties: Excel. ChartFont)](/javascript/api/excel/excel.chartfont#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартфонтупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartfont#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[size](/javascript/api/excel/excel.chartfont#size)|Размер шрифта (например, 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Чартундерлинестиле.|
|[Чартфонтдата](/javascript/api/excel/excel.chartfontdata)|[bold](/javascript/api/excel/excel.chartfontdata#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.chartfontdata#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.chartfontdata#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.chartfontdata#name)|Имя шрифта (например, Calibri)|
||[size](/javascript/api/excel/excel.chartfontdata#size)|Размер шрифта (например, 11)|
||[underline](/javascript/api/excel/excel.chartfontdata#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Чартундерлинестиле.|
|[Чартфонтлоадоптионс](/javascript/api/excel/excel.chartfontloadoptions)|[$all](/javascript/api/excel/excel.chartfontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.chartfontloadoptions#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.chartfontloadoptions#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.chartfontloadoptions#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.chartfontloadoptions#name)|Имя шрифта (например, Calibri)|
||[size](/javascript/api/excel/excel.chartfontloadoptions#size)|Размер шрифта (например, 11)|
||[underline](/javascript/api/excel/excel.chartfontloadoptions#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Чартундерлинестиле.|
|[Чартфонтупдатедата](/javascript/api/excel/excel.chartfontupdatedata)|[bold](/javascript/api/excel/excel.chartfontupdatedata#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.chartfontupdatedata#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.chartfontupdatedata#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.chartfontupdatedata#name)|Имя шрифта (например, Calibri)|
||[size](/javascript/api/excel/excel.chartfontupdatedata#size)|Размер шрифта (например, 11)|
||[underline](/javascript/api/excel/excel.chartfontupdatedata#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Чартундерлинестиле.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Представляет форматирование линий сетки диаграммы. Только для чтения.|
||[Set (Properties: Excel. ChartGridlines)](/javascript/api/excel/excel.chartgridlines#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартгридлинесупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartgridlines#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Логическое значение, определяющее, отображаются ли линии сетки оси.|
|[Чартгридлинесдата](/javascript/api/excel/excel.chartgridlinesdata)|[format](/javascript/api/excel/excel.chartgridlinesdata#format)|Представляет форматирование линий сетки диаграммы. Только для чтения.|
||[visible](/javascript/api/excel/excel.chartgridlinesdata#visible)|Логическое значение, определяющее, отображаются ли линии сетки оси.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Представляет форматирование линий диаграммы. Только для чтения.|
||[Set (Properties: Excel. ChartGridlinesFormat)](/javascript/api/excel/excel.chartgridlinesformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартгридлинесформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartgridlinesformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартгридлинесформатдата](/javascript/api/excel/excel.chartgridlinesformatdata)|[line](/javascript/api/excel/excel.chartgridlinesformatdata#line)|Представляет форматирование линий диаграммы. Только для чтения.|
|[Чартгридлинесформатлоадоптионс](/javascript/api/excel/excel.chartgridlinesformatloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartgridlinesformatloadoptions#line)|Представляет форматирование линий диаграммы.|
|[Чартгридлинесформатупдатедата](/javascript/api/excel/excel.chartgridlinesformatupdatedata)|[line](/javascript/api/excel/excel.chartgridlinesformatupdatedata#line)|Представляет форматирование линий диаграммы.|
|[Чартгридлинеслоадоптионс](/javascript/api/excel/excel.chartgridlinesloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartgridlinesloadoptions#format)|Представляет форматирование линий сетки диаграммы.|
||[visible](/javascript/api/excel/excel.chartgridlinesloadoptions#visible)|Логическое значение, определяющее, отображаются ли линии сетки оси.|
|[Чартгридлинесупдатедата](/javascript/api/excel/excel.chartgridlinesupdatedata)|[format](/javascript/api/excel/excel.chartgridlinesupdatedata#format)|Представляет форматирование линий сетки диаграммы.|
||[visible](/javascript/api/excel/excel.chartgridlinesupdatedata#visible)|Логическое значение, определяющее, отображаются ли линии сетки оси.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[накладывающиеся](/javascript/api/excel/excel.chartlegend#overlay)|Логическое значение, определяющее, должна ли легенда диаграммы перекрываться с основной частью диаграммы.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Представляет расположение легенды на диаграмме. Дополнительные сведения см. в статье Excel. Чартлежендпоситион.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Представляет форматирование легенды диаграммы, включая заливку и шрифт. Только для чтения.|
||[Set (Properties: Excel. ChartLegend)](/javascript/api/excel/excel.chartlegend#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартлежендупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartlegend#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Логическое значение, представляющее видимость объекта ChartLegend.|
|[Чартлеженддата](/javascript/api/excel/excel.chartlegenddata)|[format](/javascript/api/excel/excel.chartlegenddata#format)|Представляет форматирование легенды диаграммы, включая заливку и шрифт. Только для чтения.|
||[накладывающиеся](/javascript/api/excel/excel.chartlegenddata#overlay)|Логическое значение, определяющее, должна ли легенда диаграммы перекрываться с основной частью диаграммы.|
||[position](/javascript/api/excel/excel.chartlegenddata#position)|Представляет расположение легенды на диаграмме. Дополнительные сведения см. в статье Excel. Чартлежендпоситион.|
||[visible](/javascript/api/excel/excel.chartlegenddata#visible)|Логическое значение, представляющее видимость объекта ChartLegend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона. Только для чтения.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д., в условных обозначениях диаграммы. Только для чтения.|
||[Set (Properties: Excel. ChartLegendFormat)](/javascript/api/excel/excel.chartlegendformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартлежендформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartlegendformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартлежендформатдата](/javascript/api/excel/excel.chartlegendformatdata)|[font](/javascript/api/excel/excel.chartlegendformatdata#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д., в условных обозначениях диаграммы. Только для чтения.|
|[Чартлежендформатлоадоптионс](/javascript/api/excel/excel.chartlegendformatloadoptions)|[$all](/javascript/api/excel/excel.chartlegendformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartlegendformatloadoptions#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д., в условных обозначениях диаграммы.|
|[Чартлежендформатупдатедата](/javascript/api/excel/excel.chartlegendformatupdatedata)|[font](/javascript/api/excel/excel.chartlegendformatupdatedata#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д., в условных обозначениях диаграммы.|
|[Чартлежендлоадоптионс](/javascript/api/excel/excel.chartlegendloadoptions)|[$all](/javascript/api/excel/excel.chartlegendloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartlegendloadoptions#format)|Представляет форматирование легенды диаграммы, включая заливку и шрифт.|
||[накладывающиеся](/javascript/api/excel/excel.chartlegendloadoptions#overlay)|Логическое значение, определяющее, должна ли легенда диаграммы перекрываться с основной частью диаграммы.|
||[position](/javascript/api/excel/excel.chartlegendloadoptions#position)|Представляет расположение легенды на диаграмме. Дополнительные сведения см. в статье Excel. Чартлежендпоситион.|
||[visible](/javascript/api/excel/excel.chartlegendloadoptions#visible)|Логическое значение, представляющее видимость объекта ChartLegend.|
|[Чартлежендупдатедата](/javascript/api/excel/excel.chartlegendupdatedata)|[format](/javascript/api/excel/excel.chartlegendupdatedata#format)|Представляет форматирование легенды диаграммы, включая заливку и шрифт.|
||[накладывающиеся](/javascript/api/excel/excel.chartlegendupdatedata#overlay)|Логическое значение, определяющее, должна ли легенда диаграммы перекрываться с основной частью диаграммы.|
||[position](/javascript/api/excel/excel.chartlegendupdatedata#position)|Представляет расположение легенды на диаграмме. Дополнительные сведения см. в статье Excel. Чартлежендпоситион.|
||[visible](/javascript/api/excel/excel.chartlegendupdatedata#visible)|Логическое значение, представляющее видимость объекта ChartLegend.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Очищает формат линий элемента диаграммы.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|HTML-код цвета, представляющий цвет линий в диаграмме.|
||[Set (Properties: Excel. ChartLineFormat)](/javascript/api/excel/excel.chartlineformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартлинеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartlineformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартлинеформатдата](/javascript/api/excel/excel.chartlineformatdata)|[color](/javascript/api/excel/excel.chartlineformatdata#color)|HTML-код цвета, представляющий цвет линий в диаграмме.|
|[Чартлинеформатлоадоптионс](/javascript/api/excel/excel.chartlineformatloadoptions)|[$all](/javascript/api/excel/excel.chartlineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartlineformatloadoptions#color)|HTML-код цвета, представляющий цвет линий в диаграмме.|
|[Чартлинеформатупдатедата](/javascript/api/excel/excel.chartlineformatupdatedata)|[color](/javascript/api/excel/excel.chartlineformatupdatedata#color)|HTML-код цвета, представляющий цвет линий в диаграмме.|
|[Чартлоадоптионс](/javascript/api/excel/excel.chartloadoptions)|[$all](/javascript/api/excel/excel.chartloadoptions#$all)||
||[Axes](/javascript/api/excel/excel.chartloadoptions#axes)|Представляет оси диаграммы.|
||[dataLabels](/javascript/api/excel/excel.chartloadoptions#datalabels)|Представляет метки данных на диаграмме.|
||[format](/javascript/api/excel/excel.chartloadoptions#format)|Инкапсулирует свойства формата для области диаграммы.|
||[height](/javascript/api/excel/excel.chartloadoptions#height)|Обозначает высоту объекта диаграммы (в пунктах).|
||[left](/javascript/api/excel/excel.chartloadoptions#left)|Расстояние в пунктах от левого края диаграммы до начала листа.|
||[списком](/javascript/api/excel/excel.chartloadoptions#legend)|Представляет условные обозначения для диаграммы.|
||[name](/javascript/api/excel/excel.chartloadoptions#name)|Обозначает имя объекта диаграммы.|
||[series](/javascript/api/excel/excel.chartloadoptions#series)|Представляет один ряд данных или коллекцию рядов данных в диаграмме.|
||[заголовок](/javascript/api/excel/excel.chartloadoptions#title)|Представляет заголовок указанной диаграммы, включая его текст, видимость, положение и форматирование.|
||[top](/javascript/api/excel/excel.chartloadoptions#top)|Представляет расстояние в пунктах от верхнего края объекта до верхнего края первой строки (на листе) или до верхнего края области диаграммы (на диаграмме).|
||[width](/javascript/api/excel/excel.chartloadoptions#width)|Представляет ширину объекта диаграммы (в пунктах).|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Инкапсулирует свойства формата точки диаграммы. Только для чтения.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Возвращает значение точки диаграммы. Только для чтения.|
||[Set (Properties: Excel. ChartPoint)](/javascript/api/excel/excel.chartpoint#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартпоинтупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartpoint#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартпоинтдата](/javascript/api/excel/excel.chartpointdata)|[format](/javascript/api/excel/excel.chartpointdata#format)|Инкапсулирует свойства формата точки диаграммы. Только для чтения.|
||[value](/javascript/api/excel/excel.chartpointdata#value)|Возвращает значение точки диаграммы. Только для чтения.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Представляет формат заливки диаграммы, включающий сведения о форматировании фона. Только для чтения.|
||[Set (Properties: Excel. ChartPointFormat)](/javascript/api/excel/excel.chartpointformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартпоинтформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartpointformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартпоинтформатлоадоптионс](/javascript/api/excel/excel.chartpointformatloadoptions)|[$all](/javascript/api/excel/excel.chartpointformatloadoptions#$all)||
|[Чартпоинтлоадоптионс](/javascript/api/excel/excel.chartpointloadoptions)|[$all](/javascript/api/excel/excel.chartpointloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointloadoptions#format)|Инкапсулирует свойства формата точки диаграммы.|
||[value](/javascript/api/excel/excel.chartpointloadoptions#value)|Возвращает значение точки диаграммы. Только для чтения.|
|[Чартпоинтупдатедата](/javascript/api/excel/excel.chartpointupdatedata)|[format](/javascript/api/excel/excel.chartpointupdatedata#format)|Инкапсулирует свойства формата точки диаграммы.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Получение точки на основании ее положения в ряду.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Возвращает количество точек диаграммы в ряду. Только для чтения.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Чартпоинтсколлектионлоадоптионс](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[$all](/javascript/api/excel/excel.chartpointscollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointscollectionloadoptions#format)|Для каждого элемента в коллекции: инкапсулирует точку диаграммы свойств формата.|
||[value](/javascript/api/excel/excel.chartpointscollectionloadoptions#value)|Для каждого элемента в коллекции: Возвращает значение точки диаграммы. Только для чтения.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Представляет имя ряда в диаграмме.|
||[format](/javascript/api/excel/excel.chartseries#format)|Представляет форматирование ряда диаграммы, включая формат заливки и линий. Только для чтения.|
||[этапах](/javascript/api/excel/excel.chartseries#points)|Представляет коллекцию всех точек в ряду. Только для чтения.|
||[Set (Properties: Excel. ChartSeries)](/javascript/api/excel/excel.chartseries#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартсериесупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartseries#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Возвращает ряд в зависимости от его позиции в коллекции.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Возвращает число рядов в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Чартсериесколлектионлоадоптионс](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[$all](/javascript/api/excel/excel.chartseriescollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriescollectionloadoptions#format)|Для каждого элемента в коллекции: представляет форматирование ряда диаграммы, включающее форматирование линий и заливки.|
||[name](/javascript/api/excel/excel.chartseriescollectionloadoptions#name)|Для каждого элемента в коллекции: представляет имя ряда в диаграмме.|
||[этапах](/javascript/api/excel/excel.chartseriescollectionloadoptions#points)|Для каждого элемента в коллекции: представляет коллекцию всех точек в ряду.|
|[Чартсериесдата](/javascript/api/excel/excel.chartseriesdata)|[format](/javascript/api/excel/excel.chartseriesdata#format)|Представляет форматирование ряда диаграммы, включая формат заливки и линий. Только для чтения.|
||[name](/javascript/api/excel/excel.chartseriesdata#name)|Представляет имя ряда в диаграмме.|
||[этапах](/javascript/api/excel/excel.chartseriesdata#points)|Представляет коллекцию всех точек в ряду. Только для чтения.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Представляет формат заливки ряда диаграммы, включая сведения о форматировании фона. Только для чтения.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Представляет форматирование линий. Только для чтения.|
||[Set (Properties: Excel. ChartSeriesFormat)](/javascript/api/excel/excel.chartseriesformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартсериесформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartseriesformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартсериесформатдата](/javascript/api/excel/excel.chartseriesformatdata)|[line](/javascript/api/excel/excel.chartseriesformatdata#line)|Представляет форматирование линий. Только для чтения.|
|[Чартсериесформатлоадоптионс](/javascript/api/excel/excel.chartseriesformatloadoptions)|[$all](/javascript/api/excel/excel.chartseriesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartseriesformatloadoptions#line)|Представляет форматирование линий.|
|[Чартсериесформатупдатедата](/javascript/api/excel/excel.chartseriesformatupdatedata)|[line](/javascript/api/excel/excel.chartseriesformatupdatedata#line)|Представляет форматирование линий.|
|[Чартсериеслоадоптионс](/javascript/api/excel/excel.chartseriesloadoptions)|[$all](/javascript/api/excel/excel.chartseriesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriesloadoptions#format)|Представляет форматирование ряда диаграммы, включая формат заливки и линий.|
||[name](/javascript/api/excel/excel.chartseriesloadoptions#name)|Представляет имя ряда в диаграмме.|
||[этапах](/javascript/api/excel/excel.chartseriesloadoptions#points)|Представляет коллекцию всех точек в ряду.|
|[Чартсериесупдатедата](/javascript/api/excel/excel.chartseriesupdatedata)|[format](/javascript/api/excel/excel.chartseriesupdatedata#format)|Представляет форматирование ряда диаграммы, включая формат заливки и линий.|
||[name](/javascript/api/excel/excel.chartseriesupdatedata#name)|Представляет имя ряда в диаграмме.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[накладывающиеся](/javascript/api/excel/excel.charttitle#overlay)|Логическое значение, определяющее, отображается ли заголовок диаграммы поверх нее.|
||[format](/javascript/api/excel/excel.charttitle#format)|Представляет форматирование названия диаграммы, включая формат заливки и шрифта. Только для чтения.|
||[Set (Properties: Excel. ChartTitle)](/javascript/api/excel/excel.charttitle#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чарттитлеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.charttitle#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[text](/javascript/api/excel/excel.charttitle#text)|Представляет текст заголовка диаграммы.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Логическое значение, представляющее видимость объекта заголовка диаграммы.|
|[Чарттитледата](/javascript/api/excel/excel.charttitledata)|[format](/javascript/api/excel/excel.charttitledata#format)|Представляет форматирование названия диаграммы, включая формат заливки и шрифта. Только для чтения.|
||[накладывающиеся](/javascript/api/excel/excel.charttitledata#overlay)|Логическое значение, определяющее, отображается ли заголовок диаграммы поверх нее.|
||[text](/javascript/api/excel/excel.charttitledata#text)|Представляет текст заголовка диаграммы.|
||[visible](/javascript/api/excel/excel.charttitledata#visible)|Логическое значение, представляющее видимость объекта заголовка диаграммы.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона. Только для чтения.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Представляет атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для объекта. Только для чтения.|
||[Set (Properties: Excel. ChartTitleFormat)](/javascript/api/excel/excel.charttitleformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чарттитлеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.charttitleformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чарттитлеформатдата](/javascript/api/excel/excel.charttitleformatdata)|[font](/javascript/api/excel/excel.charttitleformatdata#font)|Представляет атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для объекта. Только для чтения.|
|[Чарттитлеформатлоадоптионс](/javascript/api/excel/excel.charttitleformatloadoptions)|[$all](/javascript/api/excel/excel.charttitleformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.charttitleformatloadoptions#font)|Представляет атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для объекта.|
|[Чарттитлеформатупдатедата](/javascript/api/excel/excel.charttitleformatupdatedata)|[font](/javascript/api/excel/excel.charttitleformatupdatedata#font)|Представляет атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для объекта.|
|[Чарттитлелоадоптионс](/javascript/api/excel/excel.charttitleloadoptions)|[$all](/javascript/api/excel/excel.charttitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttitleloadoptions#format)|Представляет форматирование названия диаграммы, включая формат заливки и шрифта.|
||[накладывающиеся](/javascript/api/excel/excel.charttitleloadoptions#overlay)|Логическое значение, определяющее, отображается ли заголовок диаграммы поверх нее.|
||[text](/javascript/api/excel/excel.charttitleloadoptions#text)|Представляет текст заголовка диаграммы.|
||[visible](/javascript/api/excel/excel.charttitleloadoptions#visible)|Логическое значение, представляющее видимость объекта заголовка диаграммы.|
|[Чарттитлеупдатедата](/javascript/api/excel/excel.charttitleupdatedata)|[format](/javascript/api/excel/excel.charttitleupdatedata#format)|Представляет форматирование названия диаграммы, включая формат заливки и шрифта.|
||[накладывающиеся](/javascript/api/excel/excel.charttitleupdatedata#overlay)|Логическое значение, определяющее, отображается ли заголовок диаграммы поверх нее.|
||[text](/javascript/api/excel/excel.charttitleupdatedata#text)|Представляет текст заголовка диаграммы.|
||[visible](/javascript/api/excel/excel.charttitleupdatedata#visible)|Логическое значение, представляющее видимость объекта заголовка диаграммы.|
|[Чартупдатедата](/javascript/api/excel/excel.chartupdatedata)|[Axes](/javascript/api/excel/excel.chartupdatedata#axes)|Представляет оси диаграммы.|
||[dataLabels](/javascript/api/excel/excel.chartupdatedata#datalabels)|Представляет метки данных на диаграмме.|
||[format](/javascript/api/excel/excel.chartupdatedata#format)|Инкапсулирует свойства формата для области диаграммы.|
||[height](/javascript/api/excel/excel.chartupdatedata#height)|Обозначает высоту объекта диаграммы (в пунктах).|
||[left](/javascript/api/excel/excel.chartupdatedata#left)|Расстояние в пунктах от левого края диаграммы до начала листа.|
||[списком](/javascript/api/excel/excel.chartupdatedata#legend)|Представляет условные обозначения для диаграммы.|
||[name](/javascript/api/excel/excel.chartupdatedata#name)|Обозначает имя объекта диаграммы.|
||[заголовок](/javascript/api/excel/excel.chartupdatedata#title)|Представляет заголовок указанной диаграммы, включая его текст, видимость, положение и форматирование.|
||[top](/javascript/api/excel/excel.chartupdatedata#top)|Представляет расстояние в пунктах от верхнего края объекта до верхнего края первой строки (на листе) или до верхнего края области диаграммы (на диаграмме).|
||[width](/javascript/api/excel/excel.chartupdatedata#width)|Представляет ширину объекта диаграммы (в пунктах).|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Возвращает объект диапазона, связанный с именем. Выдает ошибку, если именованный элемент не является диапазоном.|
||[name](/javascript/api/excel/excel.nameditem#name)|Имя объекта. Только для чтения.|
||[type](/javascript/api/excel/excel.nameditem#type)|Указывает тип значения, возвращаемый формулой имени. Дополнительные сведения см. в статье Excel. Намедитемтипе. Только для чтения.|
||[value](/javascript/api/excel/excel.nameditem#value)|Представляет значение, вычисленное по формуле имени. Если задан именованный диапазон, возвращается адрес диапазона. Только для чтения.|
||[Set (Properties: Excel. NamedItem)](/javascript/api/excel/excel.nameditem#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Намедитемупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.nameditem#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Определяет, является ли объект видимым.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Возвращает объект NamedItem, используя его имя.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Намедитемколлектионлоадоптионс](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[$all](/javascript/api/excel/excel.nameditemcollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemcollectionloadoptions#name)|Для каждого элемента в коллекции: имя объекта. Только для чтения.|
||[type](/javascript/api/excel/excel.nameditemcollectionloadoptions#type)|Для каждого элемента в коллекции: указывает тип значения, возвращаемого формулой имени. Дополнительные сведения см. в статье Excel. Намедитемтипе. Только для чтения.|
||[value](/javascript/api/excel/excel.nameditemcollectionloadoptions#value)|Для каждого элемента в коллекции: представляет значение, вычисленное с помощью формулы имени. Если задан именованный диапазон, возвращается адрес диапазона. Только для чтения.|
||[visible](/javascript/api/excel/excel.nameditemcollectionloadoptions#visible)|Для каждого элемента в коллекции: указывает, является ли объект видимым.|
|[Намедитемдата](/javascript/api/excel/excel.nameditemdata)|[name](/javascript/api/excel/excel.nameditemdata#name)|Имя объекта. Только для чтения.|
||[type](/javascript/api/excel/excel.nameditemdata#type)|Указывает тип значения, возвращаемый формулой имени. Дополнительные сведения см. в статье Excel. Намедитемтипе. Только для чтения.|
||[value](/javascript/api/excel/excel.nameditemdata#value)|Представляет значение, вычисленное по формуле имени. Если задан именованный диапазон, возвращается адрес диапазона. Только для чтения.|
||[visible](/javascript/api/excel/excel.nameditemdata#visible)|Определяет, является ли объект видимым.|
|[Намедитемлоадоптионс](/javascript/api/excel/excel.nameditemloadoptions)|[$all](/javascript/api/excel/excel.nameditemloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemloadoptions#name)|Имя объекта. Только для чтения.|
||[type](/javascript/api/excel/excel.nameditemloadoptions#type)|Указывает тип значения, возвращаемый формулой имени. Дополнительные сведения см. в статье Excel. Намедитемтипе. Только для чтения.|
||[value](/javascript/api/excel/excel.nameditemloadoptions#value)|Представляет значение, вычисленное по формуле имени. Если задан именованный диапазон, возвращается адрес диапазона. Только для чтения.|
||[visible](/javascript/api/excel/excel.nameditemloadoptions#visible)|Определяет, является ли объект видимым.|
|[Намедитемупдатедата](/javascript/api/excel/excel.nameditemupdatedata)|[visible](/javascript/api/excel/excel.nameditemupdatedata#visible)|Определяет, является ли объект видимым.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.range#clear-applyto-)|Очищает значения, формат, заливку, границу диапазона и т. д.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Очищает значения, формат, заливку, границу диапазона и т. д.|
||[Delete (Shift: "Up" \| "Left")](/javascript/api/excel/excel.range#delete-shift-)|Удаляет ячейки, связанные с диапазоном.|
||[Delete (Shift: Excel. Делетешифтдиректион)](/javascript/api/excel/excel.range#delete-shift-)|Удаляет ячейки, связанные с диапазоном.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[getBoundingRect (anotherRange: строка \| Range)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Возвращает наименьший объект диапазона, включающий в себя заданные диапазоны. Например, GetBoundingRect для "B2:C5" и "D10:E15" возвращает значение "B2:E15".|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца. Ячейка может находиться вне границ родительского диапазона, пока она остается в сетке листа. Возвращаемая ячейка располагается относительно верхней левой ячейки диапазона.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Возвращает столбец в диапазоне.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Получает объект, представляющий весь столбец диапазона (например, если текущий диапазон представляет ячейки "B4: E11", а `getEntireColumn` — диапазон, представляющий столбцы "б:е").|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Получает объект, представляющий всю строку диапазона (например, если текущий диапазон представляет ячейки "B4: E11", а `GetEntireRow` — диапазон, представляющий строки "4:11").|
||[пересечение (anotherRange: строка \| Range)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Возвращает объект диапазона, представляющий собой прямоугольное пересечение заданных диапазонов.|
||[Жетластцелл ()](/javascript/api/excel/excel.range#getlastcell--)|Возвращает последнюю ячейку в диапазоне. Например, последняя ячейка диапазона B2:D5 — D5.|
||[Жетластколумн ()](/javascript/api/excel/excel.range#getlastcolumn--)|Возвращает последний столбец в диапазоне. Например, последний столбец диапазона B2:D5 — D2:D5.|
||[Жетластров ()](/javascript/api/excel/excel.range#getlastrow--)|Возвращает последнюю строку в диапазоне. Например, последняя строка в диапазоне "B2:D5" — "B5:D5".|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Возвращает объект, представляющий диапазон, который смещен от указанного диапазона. Измерение возвращаемого диапазона будет соответствовать этому диапазону. Если результирующий диапазон выходит за пределы таблицы листа, возникнет ошибка.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Возвращает строку из диапазона.|
||[INSERT (Shift: "Down" \| "" справа ")](/javascript/api/excel/excel.range#insert-shift-)|Вставляет ячейку или диапазон ячеек на лист вместо этого диапазона, а также сдвигает другие ячейки, чтобы освободить место. Возвращает новый объект Range в пустом месте.|
||[INSERT (Shift: Excel. Инсертшифтдиректион)](/javascript/api/excel/excel.range#insert-shift-)|Вставляет ячейку или диапазон ячеек на лист вместо этого диапазона, а также сдвигает другие ячейки, чтобы освободить место. Возвращает новый объект Range в пустом месте.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Представляет код числового формата Excel для заданного диапазона.|
||[address](/javascript/api/excel/excel.range#address)|Представляет ссылку на диапазон в стиле A1. Значение Address будет содержать ссылку на лист (например, "Лист1! A1: B4). Только для чтения.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Представляет ссылку на указанный диапазон на языке пользователя. Только для чтения.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Количество ячеек в диапазоне. Этот API возвращает значение -1, если количество ячеек превышает 2^31-1 (2,147,483,647). Только для чтения.|
||[Число](/javascript/api/excel/excel.range#columncount)|Представляет общее количество столбцов в диапазоне. Только для чтения.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Представляет номер столбца первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
||[format](/javascript/api/excel/excel.range#format)|Возвращает объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона. Только для чтения.|
||[Стро](/javascript/api/excel/excel.range#rowcount)|Возвращает общее количество строк в диапазоне. Только для чтения.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Возвращает номер строки первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
||[text](/javascript/api/excel/excel.range#text)|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Представляет тип данных каждой ячейки. Только для чтения.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|Лист, содержащий текущий диапазон. Только для чтения.|
||[select()](/javascript/api/excel/excel.range#select--)|Выбирает указанный диапазон в пользовательском интерфейсе Excel.|
||[Set (Properties: Excel. Range)](/javascript/api/excel/excel.range#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Ранжеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.range#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[track()](/javascript/api/excel/excel.range#track--)|Отслеживает объект для автоматической корректировки с учетом окружающих изменений в документе. Этот вызов является сокращением для context.trackedObjects.add(thisObject). Если этот объект используется в вызовах .sync и вне последовательного выполнения пакета .run с возникновением ошибки InvalidObjectPath при установке свойства или вызове метода для объекта, необходимо было добавить объект в коллекцию отслеживаемых объектов при первоначальном создании объекта.|
||[untrack()](/javascript/api/excel/excel.range#untrack--)|Освобождает память, связанную с этим объектом, если он ранее отслеживался. Этот вызов является сокращением для context.trackedObjects.remove(thisObject). Наличие большого количества отслеживаемых объектов замедляет ведущее приложение, поэтому не забывайте освобождать любые добавленные объекты после завершения их использования. Перед фактическим освобождением памяти потребуется вызвать метод context.sync().|
||[values](/javascript/api/excel/excel.range#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Сидеиндекс](/javascript/api/excel/excel.rangeborder#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Бордериндекс. Только для чтения.|
||[Set (Properties: Excel. RangeBorder)](/javascript/api/excel/excel.rangeborder#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Ранжебордерупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.rangeborder#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[style](/javascript/api/excel/excel.rangeborder#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Определяет толщину границы вокруг диапазона. Дополнительные сведения см. в статье Excel. Бордервеигхт.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[GetItem \| (index: "еджетоп" "еджеботтом" \| "еджелефт" \| "еджеригхт" \| "инсидевертикал" \| "инсидехоризонтал" \| "диагоналдовн" \| "DiagonalUp")](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[GetItem (index: Excel. Бордериндекс)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Возвращает объект границы по его индексу.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Количество объектов границы в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Ранжебордерколлектионлоадоптионс](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.rangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangebordercollectionloadoptions#color)|Для каждого элемента в коллекции: HTML-код цвета, представляющий цвет линии границы, формы #RRGGBB (например, "FFA500") или в виде именованного цвета HTML (например, "Апельсин").|
||[Сидеиндекс](/javascript/api/excel/excel.rangebordercollectionloadoptions#sideindex)|Для каждого элемента в коллекции: значение константы, которое указывает на конкретную сторону границы. Дополнительные сведения см. в статье Excel. Бордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.rangebordercollectionloadoptions#style)|Для каждого элемента в коллекции: одна из констант стиля линии, определяющая стиль линии для границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
||[weight](/javascript/api/excel/excel.rangebordercollectionloadoptions#weight)|Для каждого элемента в коллекции: определяет толщину границы вокруг диапазона. Дополнительные сведения см. в статье Excel. Бордервеигхт.|
|[Ранжебордердата](/javascript/api/excel/excel.rangeborderdata)|[color](/javascript/api/excel/excel.rangeborderdata#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Сидеиндекс](/javascript/api/excel/excel.rangeborderdata#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Бордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.rangeborderdata#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
||[weight](/javascript/api/excel/excel.rangeborderdata#weight)|Определяет толщину границы вокруг диапазона. Дополнительные сведения см. в статье Excel. Бордервеигхт.|
|[Ранжебордерлоадоптионс](/javascript/api/excel/excel.rangeborderloadoptions)|[$all](/javascript/api/excel/excel.rangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangeborderloadoptions#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Сидеиндекс](/javascript/api/excel/excel.rangeborderloadoptions#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Бордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.rangeborderloadoptions#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
||[weight](/javascript/api/excel/excel.rangeborderloadoptions#weight)|Определяет толщину границы вокруг диапазона. Дополнительные сведения см. в статье Excel. Бордервеигхт.|
|[Ранжебордерупдатедата](/javascript/api/excel/excel.rangeborderupdatedata)|[color](/javascript/api/excel/excel.rangeborderupdatedata#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[style](/javascript/api/excel/excel.rangeborderupdatedata#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
||[weight](/javascript/api/excel/excel.rangeborderupdatedata#weight)|Определяет толщину границы вокруг диапазона. Дополнительные сведения см. в статье Excel. Бордервеигхт.|
|[Ранжедата](/javascript/api/excel/excel.rangedata)|[address](/javascript/api/excel/excel.rangedata#address)|Представляет ссылку на диапазон в стиле A1. Значение Address будет содержать ссылку на лист (например, "Лист1! A1: B4). Только для чтения.|
||[addressLocal](/javascript/api/excel/excel.rangedata#addresslocal)|Представляет ссылку на указанный диапазон на языке пользователя. Только для чтения.|
||[cellCount](/javascript/api/excel/excel.rangedata#cellcount)|Количество ячеек в диапазоне. Этот API возвращает значение -1, если количество ячеек превышает 2^31-1 (2,147,483,647). Только для чтения.|
||[Число](/javascript/api/excel/excel.rangedata#columncount)|Представляет общее количество столбцов в диапазоне. Только для чтения.|
||[columnIndex](/javascript/api/excel/excel.rangedata#columnindex)|Представляет номер столбца первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
||[format](/javascript/api/excel/excel.rangedata#format)|Возвращает объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона. Только для чтения.|
||[formulas](/javascript/api/excel/excel.rangedata#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangedata#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[numberFormat](/javascript/api/excel/excel.rangedata#numberformat)|Представляет код числового формата Excel для заданного диапазона.|
||[Стро](/javascript/api/excel/excel.rangedata#rowcount)|Возвращает общее количество строк в диапазоне. Только для чтения.|
||[rowIndex](/javascript/api/excel/excel.rangedata#rowindex)|Возвращает номер строки первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
||[text](/javascript/api/excel/excel.rangedata#text)|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.rangedata#valuetypes)|Представляет тип данных каждой ячейки. Только для чтения.|
||[values](/javascript/api/excel/excel.rangedata#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Сброс фона диапазона.|
||[color](/javascript/api/excel/excel.rangefill#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова в HTML (например, orange).|
||[Set (Properties: Excel. RangeFill)](/javascript/api/excel/excel.rangefill#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Ранжефиллупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.rangefill#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Ранжефиллдата](/javascript/api/excel/excel.rangefilldata)|[color](/javascript/api/excel/excel.rangefilldata#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова в HTML (например, orange).|
|[Ранжефилллоадоптионс](/javascript/api/excel/excel.rangefillloadoptions)|[$all](/javascript/api/excel/excel.rangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangefillloadoptions#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова в HTML (например, orange).|
|[Ранжефиллупдатедата](/javascript/api/excel/excel.rangefillupdatedata)|[color](/javascript/api/excel/excel.rangefillupdatedata#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова в HTML (например, orange).|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.rangefont#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.rangefont#name)|Имя шрифта (например, Calibri)|
||[Set (Properties: Excel. RangeFont)](/javascript/api/excel/excel.rangefont#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Ранжефонтупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.rangefont#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[size](/javascript/api/excel/excel.rangefont#size)|размер шрифта|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Ранжеундерлинестиле.|
|[Ранжефонтдата](/javascript/api/excel/excel.rangefontdata)|[bold](/javascript/api/excel/excel.rangefontdata#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.rangefontdata#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.rangefontdata#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.rangefontdata#name)|Имя шрифта (например, Calibri)|
||[size](/javascript/api/excel/excel.rangefontdata#size)|размер шрифта|
||[underline](/javascript/api/excel/excel.rangefontdata#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Ранжеундерлинестиле.|
|[Ранжефонтлоадоптионс](/javascript/api/excel/excel.rangefontloadoptions)|[$all](/javascript/api/excel/excel.rangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.rangefontloadoptions#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.rangefontloadoptions#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.rangefontloadoptions#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.rangefontloadoptions#name)|Имя шрифта (например, Calibri)|
||[size](/javascript/api/excel/excel.rangefontloadoptions#size)|размер шрифта|
||[underline](/javascript/api/excel/excel.rangefontloadoptions#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Ранжеундерлинестиле.|
|[Ранжефонтупдатедата](/javascript/api/excel/excel.rangefontupdatedata)|[bold](/javascript/api/excel/excel.rangefontupdatedata#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.rangefontupdatedata#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.rangefontupdatedata#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.rangefontupdatedata#name)|Имя шрифта (например, Calibri)|
||[size](/javascript/api/excel/excel.rangefontupdatedata#size)|размер шрифта|
||[underline](/javascript/api/excel/excel.rangefontupdatedata#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Ранжеундерлинестиле.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Представляет выравнивание по горизонтали для указанного объекта. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|Коллекция объектов границ, которые применяются ко всему диапазону. Только для чтения.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Возвращает объект заливки, определенный для всего диапазона. Только для чтения.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Возвращает объект шрифта, определенный для всего диапазона. Только для чтения.|
||[Set (Properties: Excel. RangeFormat)](/javascript/api/excel/excel.rangeformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Ранжеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.rangeformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Представляет выравнивание по вертикали для указанного объекта. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Указывает, использует ли Excel обтекание текстом для объекта. Значение null указывает, что для диапазона в целом не применяется согласованный параметр обтекания.|
|[Ранжеформатдата](/javascript/api/excel/excel.rangeformatdata)|[borders](/javascript/api/excel/excel.rangeformatdata#borders)|Коллекция объектов границ, которые применяются ко всему диапазону. Только для чтения.|
||[fill](/javascript/api/excel/excel.rangeformatdata#fill)|Возвращает объект заливки, определенный для всего диапазона. Только для чтения.|
||[font](/javascript/api/excel/excel.rangeformatdata#font)|Возвращает объект шрифта, определенный для всего диапазона. Только для чтения.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatdata#horizontalalignment)|Представляет выравнивание по горизонтали для указанного объекта. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatdata#verticalalignment)|Представляет выравнивание по вертикали для указанного объекта. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformatdata#wraptext)|Указывает, использует ли Excel обтекание текстом для объекта. Значение null указывает, что для диапазона в целом не применяется согласованный параметр обтекания.|
|[Ранжеформатлоадоптионс](/javascript/api/excel/excel.rangeformatloadoptions)|[$all](/javascript/api/excel/excel.rangeformatloadoptions#$all)||
||[borders](/javascript/api/excel/excel.rangeformatloadoptions#borders)|Коллекция объектов границ, которые применяются ко всему диапазону.|
||[fill](/javascript/api/excel/excel.rangeformatloadoptions#fill)|Возвращает объект заливки, определенный для всего диапазона.|
||[font](/javascript/api/excel/excel.rangeformatloadoptions#font)|Возвращает объект шрифта, определенный для всего диапазона.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#horizontalalignment)|Представляет выравнивание по горизонтали для указанного объекта. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#verticalalignment)|Представляет выравнивание по вертикали для указанного объекта. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformatloadoptions#wraptext)|Указывает, использует ли Excel обтекание текстом для объекта. Значение null указывает, что для диапазона в целом не применяется согласованный параметр обтекания.|
|[Ранжеформатупдатедата](/javascript/api/excel/excel.rangeformatupdatedata)|[borders](/javascript/api/excel/excel.rangeformatupdatedata#borders)|Коллекция объектов границ, которые применяются ко всему диапазону.|
||[fill](/javascript/api/excel/excel.rangeformatupdatedata#fill)|Возвращает объект заливки, определенный для всего диапазона.|
||[font](/javascript/api/excel/excel.rangeformatupdatedata#font)|Возвращает объект шрифта, определенный для всего диапазона.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#horizontalalignment)|Представляет выравнивание по горизонтали для указанного объекта. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#verticalalignment)|Представляет выравнивание по вертикали для указанного объекта. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformatupdatedata#wraptext)|Указывает, использует ли Excel обтекание текстом для объекта. Значение null указывает, что для диапазона в целом не применяется согласованный параметр обтекания.|
|[Ранжелоадоптионс](/javascript/api/excel/excel.rangeloadoptions)|[$all](/javascript/api/excel/excel.rangeloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeloadoptions#address)|Представляет ссылку на диапазон в стиле A1. Значение Address будет содержать ссылку на лист (например, "Лист1! A1: B4). Только для чтения.|
||[addressLocal](/javascript/api/excel/excel.rangeloadoptions#addresslocal)|Представляет ссылку на указанный диапазон на языке пользователя. Только для чтения.|
||[cellCount](/javascript/api/excel/excel.rangeloadoptions#cellcount)|Количество ячеек в диапазоне. Этот API возвращает значение -1, если количество ячеек превышает 2^31-1 (2,147,483,647). Только для чтения.|
||[Число](/javascript/api/excel/excel.rangeloadoptions#columncount)|Представляет общее количество столбцов в диапазоне. Только для чтения.|
||[columnIndex](/javascript/api/excel/excel.rangeloadoptions#columnindex)|Представляет номер столбца первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
||[format](/javascript/api/excel/excel.rangeloadoptions#format)|Возвращает объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона.|
||[formulas](/javascript/api/excel/excel.rangeloadoptions#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeloadoptions#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[numberFormat](/javascript/api/excel/excel.rangeloadoptions#numberformat)|Представляет код числового формата Excel для заданного диапазона.|
||[Стро](/javascript/api/excel/excel.rangeloadoptions#rowcount)|Возвращает общее количество строк в диапазоне. Только для чтения.|
||[rowIndex](/javascript/api/excel/excel.rangeloadoptions#rowindex)|Возвращает номер строки первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
||[text](/javascript/api/excel/excel.rangeloadoptions#text)|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.rangeloadoptions#valuetypes)|Представляет тип данных каждой ячейки. Только для чтения.|
||[values](/javascript/api/excel/excel.rangeloadoptions#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
||[worksheet](/javascript/api/excel/excel.rangeloadoptions#worksheet)|Лист, содержащий текущий диапазон.|
|[Ранжеупдатедата](/javascript/api/excel/excel.rangeupdatedata)|[format](/javascript/api/excel/excel.rangeupdatedata#format)|Возвращает объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона.|
||[formulas](/javascript/api/excel/excel.rangeupdatedata#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeupdatedata#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[numberFormat](/javascript/api/excel/excel.rangeupdatedata#numberformat)|Представляет код числового формата Excel для заданного диапазона.|
||[values](/javascript/api/excel/excel.rangeupdatedata#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Удаляет таблицу.|
||[Жетдатабодиранже ()](/javascript/api/excel/excel.table#getdatabodyrange--)|Получает объект диапазона, связанный с телом данных таблицы.|
||[Жесеадерровранже ()](/javascript/api/excel/excel.table#getheaderrowrange--)|Получает объект диапазона, связанный со строкой заголовков таблицы.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Получает объект диапазона, связанный со всей таблицей.|
||[Жеттоталровранже ()](/javascript/api/excel/excel.table#gettotalrowrange--)|Получает объект диапазона, связанный со строкой итогов таблицы.|
||[name](/javascript/api/excel/excel.table#name)|Имя таблицы.|
||[столбцы](/javascript/api/excel/excel.table#columns)|Представляет коллекцию всех столбцов в таблице. Только для чтения.|
||[id](/javascript/api/excel/excel.table#id)|Возвращает значение, однозначно идентифицирующее таблицу в данной книге. Значение идентификатора остается прежним, даже если переименовать таблицу. Только для чтения.|
||[строки](/javascript/api/excel/excel.table#rows)|Представляет коллекцию всех строк в таблице. Только для чтения.|
||[Set (Properties: Excel. Table)](/javascript/api/excel/excel.table#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Таблеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.table#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[Шовхеадерс](/javascript/api/excel/excel.table#showheaders)|Указывает, отображается ли строка заголовков. Можно задать это значение, чтобы отобразить или скрыть строку заголовков.|
||[Шовтоталс](/javascript/api/excel/excel.table#showtotals)|Указывает, отображается ли строка итогов. Можно задать это значение, чтобы отобразить или скрыть строку итогов.|
||[style](/javascript/api/excel/excel.table#style)|Постоянное значение, представляющее стиль таблицы. Возможные значения: от TableStyleLight1 до TableStyleLight21, от TableStyleMedium1 до TableStyleMedium28, от TableStyleStyleDark1 до TableStyleStyleDark11. Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[Add (Address: строка \| диапазона, hasHeaders: Boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Создание таблицы. Объект или исходный адрес диапазона определяет лист, на который будет добавлена таблица. Если добавить таблицу не удается (например, если адрес недействителен или одна таблица будет перекрываться другой), выводится сообщение об ошибке.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Получает таблицу по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Получает таблицу на основании ее позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Возвращает количество таблиц в книге. Только для чтения.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Таблеколлектионлоадоптионс](/javascript/api/excel/excel.tablecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecollectionloadoptions#$all)||
||[столбцы](/javascript/api/excel/excel.tablecollectionloadoptions#columns)|Для каждого элемента в коллекции: представляет коллекцию всех столбцов в таблице.|
||[id](/javascript/api/excel/excel.tablecollectionloadoptions#id)|Для каждого элемента в коллекции: Возвращает значение, однозначно идентифицирующее таблицу в заданной книге. Значение идентификатора остается прежним, даже если переименовать таблицу. Только для чтения.|
||[name](/javascript/api/excel/excel.tablecollectionloadoptions#name)|Для каждого элемента в коллекции: имя таблицы.|
||[строки](/javascript/api/excel/excel.tablecollectionloadoptions#rows)|Для каждого элемента в коллекции: представляет коллекцию всех строк в таблице.|
||[Шовхеадерс](/javascript/api/excel/excel.tablecollectionloadoptions#showheaders)|Для каждого элемента в коллекции: указывает, видима ли строка заголовков. Можно задать это значение, чтобы отобразить или скрыть строку заголовков.|
||[Шовтоталс](/javascript/api/excel/excel.tablecollectionloadoptions#showtotals)|Для каждого элемента в коллекции: указывает, видима ли строка итогов. Можно задать это значение, чтобы отобразить или скрыть строку итогов.|
||[style](/javascript/api/excel/excel.tablecollectionloadoptions#style)|Для каждого элемента в коллекции: значение константы, представляющее стиль таблицы. Возможные значения: от TableStyleLight1 до TableStyleLight21, от TableStyleMedium1 до TableStyleMedium28, от TableStyleStyleDark1 до TableStyleStyleDark11. Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Удаляет столбец из таблицы.|
||[Жетдатабодиранже ()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Получает объект диапазона, связанный с текстом данных столбца.|
||[Жесеадерровранже ()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Получает объект диапазона, связанный со строкой заголовков столбца.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Получает объект диапазона, связанный со всем столбцом.|
||[Жеттоталровранже ()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Получает объект диапазона, связанный со строкой итогов столбца.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Представляет имя столбца таблицы.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Возвращает уникальный ключ, идентифицирующий столбец в таблице. Только для чтения.|
||[индекс](/javascript/api/excel/excel.tablecolumn#index)|Возвращает номер индекса столбца в коллекции столбцов таблицы. Используется нулевой индекс. Только для чтения.|
||[Set (Properties: Excel. TableColumn)](/javascript/api/excel/excel.tablecolumn#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Таблеколумнупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.tablecolumn#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[Add (index?: число, Values?: массив<массив<логический \| номер \| строки>> \| логический \| номер \| строки, Name?: строка)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Добавляет новый столбец в таблицу.|
||[GetItem (ключ: число \| строка)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Возвращает объект column по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Возвращает столбец на основании его позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Возвращает количество столбцов в таблице. Только для чтения.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Таблеколумнколлектионлоадоптионс](/javascript/api/excel/excel.tablecolumncollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecolumncollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumncollectionloadoptions#id)|Для каждого элемента в коллекции: Возвращает уникальный ключ, который определяет столбец в таблице. Только для чтения.|
||[индекс](/javascript/api/excel/excel.tablecolumncollectionloadoptions#index)|Для каждого элемента в коллекции: Возвращает номер индекса столбца в коллекции Columns таблицы. Используется нулевой индекс. Только для чтения.|
||[name](/javascript/api/excel/excel.tablecolumncollectionloadoptions#name)|Для каждого элемента в коллекции: представляет имя столбца таблицы.|
||[values](/javascript/api/excel/excel.tablecolumncollectionloadoptions#values)|Для каждого элемента в коллекции: представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Таблеколумндата](/javascript/api/excel/excel.tablecolumndata)|[id](/javascript/api/excel/excel.tablecolumndata#id)|Возвращает уникальный ключ, идентифицирующий столбец в таблице. Только для чтения.|
||[индекс](/javascript/api/excel/excel.tablecolumndata#index)|Возвращает номер индекса столбца в коллекции столбцов таблицы. Используется нулевой индекс. Только для чтения.|
||[name](/javascript/api/excel/excel.tablecolumndata#name)|Представляет имя столбца таблицы.|
||[values](/javascript/api/excel/excel.tablecolumndata#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Таблеколумнлоадоптионс](/javascript/api/excel/excel.tablecolumnloadoptions)|[$all](/javascript/api/excel/excel.tablecolumnloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumnloadoptions#id)|Возвращает уникальный ключ, идентифицирующий столбец в таблице. Только для чтения.|
||[индекс](/javascript/api/excel/excel.tablecolumnloadoptions#index)|Возвращает номер индекса столбца в коллекции столбцов таблицы. Используется нулевой индекс. Только для чтения.|
||[name](/javascript/api/excel/excel.tablecolumnloadoptions#name)|Представляет имя столбца таблицы.|
||[values](/javascript/api/excel/excel.tablecolumnloadoptions#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Таблеколумнупдатедата](/javascript/api/excel/excel.tablecolumnupdatedata)|[name](/javascript/api/excel/excel.tablecolumnupdatedata#name)|Представляет имя столбца таблицы.|
||[values](/javascript/api/excel/excel.tablecolumnupdatedata#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[TableData](/javascript/api/excel/excel.tabledata)|[столбцы](/javascript/api/excel/excel.tabledata#columns)|Представляет коллекцию всех столбцов в таблице. Только для чтения.|
||[id](/javascript/api/excel/excel.tabledata#id)|Возвращает значение, однозначно идентифицирующее таблицу в данной книге. Значение идентификатора остается прежним, даже если переименовать таблицу. Только для чтения.|
||[name](/javascript/api/excel/excel.tabledata#name)|Имя таблицы.|
||[строки](/javascript/api/excel/excel.tabledata#rows)|Представляет коллекцию всех строк в таблице. Только для чтения.|
||[Шовхеадерс](/javascript/api/excel/excel.tabledata#showheaders)|Указывает, отображается ли строка заголовков. Можно задать это значение, чтобы отобразить или скрыть строку заголовков.|
||[Шовтоталс](/javascript/api/excel/excel.tabledata#showtotals)|Указывает, отображается ли строка итогов. Можно задать это значение, чтобы отобразить или скрыть строку итогов.|
||[style](/javascript/api/excel/excel.tabledata#style)|Постоянное значение, представляющее стиль таблицы. Возможные значения: от TableStyleLight1 до TableStyleLight21, от TableStyleMedium1 до TableStyleMedium28, от TableStyleStyleDark1 до TableStyleStyleDark11. Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
|[Таблелоадоптионс](/javascript/api/excel/excel.tableloadoptions)|[$all](/javascript/api/excel/excel.tableloadoptions#$all)||
||[столбцы](/javascript/api/excel/excel.tableloadoptions#columns)|Представляет коллекцию всех столбцов в таблице.|
||[id](/javascript/api/excel/excel.tableloadoptions#id)|Возвращает значение, однозначно идентифицирующее таблицу в данной книге. Значение идентификатора остается прежним, даже если переименовать таблицу. Только для чтения.|
||[name](/javascript/api/excel/excel.tableloadoptions#name)|Имя таблицы.|
||[строки](/javascript/api/excel/excel.tableloadoptions#rows)|Представляет коллекцию всех строк в таблице.|
||[Шовхеадерс](/javascript/api/excel/excel.tableloadoptions#showheaders)|Указывает, отображается ли строка заголовков. Можно задать это значение, чтобы отобразить или скрыть строку заголовков.|
||[Шовтоталс](/javascript/api/excel/excel.tableloadoptions#showtotals)|Указывает, отображается ли строка итогов. Можно задать это значение, чтобы отобразить или скрыть строку итогов.|
||[style](/javascript/api/excel/excel.tableloadoptions#style)|Постоянное значение, представляющее стиль таблицы. Возможные значения: от TableStyleLight1 до TableStyleLight21, от TableStyleMedium1 до TableStyleMedium28, от TableStyleStyleDark1 до TableStyleStyleDark11. Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Удаляет строку из таблицы.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Получает объект диапазона, связанный со всей строкой.|
||[индекс](/javascript/api/excel/excel.tablerow#index)|Возвращает номер индекса строки в коллекции строк таблицы. Используется нулевой индекс. Только для чтения.|
||[Set (Properties: Excel. TableRow)](/javascript/api/excel/excel.tablerow#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Таблеровупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.tablerow#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[values](/javascript/api/excel/excel.tablerow#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[Add (index?: число, Values?: массив<массив<логический \| номер \| строки>> \| логический \| номер \| строки)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Добавляет одну или несколько строк в таблицу. Возвращается объект, находящийся над новыми строками.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Получает строку на основании ее позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Возвращает количество строк в таблице. Только для чтения.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Таблеровколлектионлоадоптионс](/javascript/api/excel/excel.tablerowcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablerowcollectionloadoptions#$all)||
||[индекс](/javascript/api/excel/excel.tablerowcollectionloadoptions#index)|Для каждого элемента в коллекции: Возвращает номер индекса строки в коллекции Rows таблицы. Используется нулевой индекс. Только для чтения.|
||[values](/javascript/api/excel/excel.tablerowcollectionloadoptions#values)|Для каждого элемента в коллекции: представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Таблеровдата](/javascript/api/excel/excel.tablerowdata)|[индекс](/javascript/api/excel/excel.tablerowdata#index)|Возвращает номер индекса строки в коллекции строк таблицы. Используется нулевой индекс. Только для чтения.|
||[values](/javascript/api/excel/excel.tablerowdata#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Таблеровлоадоптионс](/javascript/api/excel/excel.tablerowloadoptions)|[$all](/javascript/api/excel/excel.tablerowloadoptions#$all)||
||[индекс](/javascript/api/excel/excel.tablerowloadoptions#index)|Возвращает номер индекса строки в коллекции строк таблицы. Используется нулевой индекс. Только для чтения.|
||[values](/javascript/api/excel/excel.tablerowloadoptions#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Таблеровупдатедата](/javascript/api/excel/excel.tablerowupdatedata)|[values](/javascript/api/excel/excel.tablerowupdatedata#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Таблеупдатедата](/javascript/api/excel/excel.tableupdatedata)|[name](/javascript/api/excel/excel.tableupdatedata#name)|Имя таблицы.|
||[showHeaders](/javascript/api/excel/excel.tableupdatedata#showheaders)|Указывает, отображается ли строка заголовков. Можно задать это значение, чтобы отобразить или скрыть строку заголовков.|
||[Шовтоталс](/javascript/api/excel/excel.tableupdatedata#showtotals)|Указывает, отображается ли строка итогов. Можно задать это значение, чтобы отобразить или скрыть строку итогов.|
||[style](/javascript/api/excel/excel.tableupdatedata#style)|Постоянное значение, представляющее стиль таблицы. Возможные значения: от TableStyleLight1 до TableStyleLight21, от TableStyleMedium1 до TableStyleMedium28, от TableStyleStyleDark1 до TableStyleStyleDark11. Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Функцией getselectedrange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|Получает текущий выделенный диапазон из книги. Если выбрано несколько диапазонов, этот метод выдаст ошибку.|
||[application](/javascript/api/excel/excel.workbook#application)|Представляет экземпляр приложения Excel, который содержит эту книгу. Только для чтения.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Представляет коллекцию привязок, включенных в книгу. Только для чтения.|
||[names](/javascript/api/excel/excel.workbook#names)|Представляет коллекцию именованных элементов в книге (именованные диапазоны и константы). Только для чтения.|
||[Table](/javascript/api/excel/excel.workbook#tables)|Представляет коллекцию таблиц, сопоставленных с книгой. Только для чтения.|
||[листов](/javascript/api/excel/excel.workbook#worksheets)|Представляет коллекцию листов, сопоставленных с книгой. Только для чтения.|
||[Set (Properties: Excel. Workbook)](/javascript/api/excel/excel.workbook#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Воркбукупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.workbook#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Воркбукдата](/javascript/api/excel/excel.workbookdata)|[bindings](/javascript/api/excel/excel.workbookdata#bindings)|Представляет коллекцию привязок, включенных в книгу. Только для чтения.|
||[names](/javascript/api/excel/excel.workbookdata#names)|Представляет коллекцию именованных элементов в книге (именованные диапазоны и константы). Только для чтения.|
||[Table](/javascript/api/excel/excel.workbookdata#tables)|Представляет коллекцию таблиц, сопоставленных с книгой. Только для чтения.|
||[листов](/javascript/api/excel/excel.workbookdata#worksheets)|Представляет коллекцию листов, сопоставленных с книгой. Только для чтения.|
|[Воркбуклоадоптионс](/javascript/api/excel/excel.workbookloadoptions)|[$all](/javascript/api/excel/excel.workbookloadoptions#$all)||
||[application](/javascript/api/excel/excel.workbookloadoptions#application)|Представляет экземпляр приложения Excel, который содержит эту книгу.|
||[bindings](/javascript/api/excel/excel.workbookloadoptions#bindings)|Представляет коллекцию привязок, включенных в книгу.|
||[Table](/javascript/api/excel/excel.workbookloadoptions#tables)|Представляет коллекцию таблиц, сопоставленных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Активация листа в пользовательском интерфейсе Excel.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Удаляет лист из книги. Обратите внимание, что если для отображения листа задано значение "Верихидден", операция удаления завершится с помощью GeneralException.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца. Ячейка может находиться вне границ родительского диапазона, пока она остается в сетке листа.|
||[GetString (Address?: строка)](/javascript/api/excel/excel.worksheet#getrange-address-)|Получает объект Range, представляющий отдельный прямоугольный блок ячеек, заданный по адресу или имени.|
||[name](/javascript/api/excel/excel.worksheet#name)|Отображаемое имя листа.|
||[position](/javascript/api/excel/excel.worksheet#position)|Положение листа (начиная с нуля) в книге.|
||[темп](/javascript/api/excel/excel.worksheet#charts)|Возвращает коллекцию диаграмм, имеющихся на листе. Только для чтения.|
||[id](/javascript/api/excel/excel.worksheet#id)|Возвращает значение, однозначно идентифицирующее лист в данной книге. Значение идентификатора остается прежним, даже если переименовать или переместить лист. Только для чтения.|
||[Table](/javascript/api/excel/excel.worksheet#tables)|Коллекция таблиц, имеющихся на листе. Только для чтения.|
||[Set (Properties: Excel. лист)](/javascript/api/excel/excel.worksheet#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Воркшитупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.worksheet#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[доступности](/javascript/api/excel/excel.worksheet#visibility)|Видимость листа.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[Добавить (имя?: строка)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Добавляет новый лист в книгу. Лист будет добавлен в конец набора имеющихся листов. Если вы хотите активировать только что добавленный лист, вызовите команду .activate().|
||[Жетактивеворкшит ()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Получает текущий активный лист в книге.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Получает объект листа по его имени или ИД.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Воркшитколлектионлоадоптионс](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[$all](/javascript/api/excel/excel.worksheetcollectionloadoptions#$all)||
||[темп](/javascript/api/excel/excel.worksheetcollectionloadoptions#charts)|Для каждого элемента в коллекции: Возвращает коллекцию диаграмм, которые являются частью листа.|
||[id](/javascript/api/excel/excel.worksheetcollectionloadoptions#id)|Для каждого элемента в коллекции: Возвращает значение, однозначно идентифицирующее лист в заданной книге. Значение идентификатора остается прежним, даже если переименовать или переместить лист. Только для чтения.|
||[name](/javascript/api/excel/excel.worksheetcollectionloadoptions#name)|Для каждого элемента в коллекции: отображаемое имя листа.|
||[position](/javascript/api/excel/excel.worksheetcollectionloadoptions#position)|Для каждого элемента в коллекции: позиция листа в книге (с отсчетом от нуля).|
||[Table](/javascript/api/excel/excel.worksheetcollectionloadoptions#tables)|Для каждого элемента в коллекции: Коллекция таблиц, которые являются частью листа.|
||[доступности](/javascript/api/excel/excel.worksheetcollectionloadoptions#visibility)|Для каждого элемента в коллекции: видимость листа.|
|[Воркшитдата](/javascript/api/excel/excel.worksheetdata)|[темп](/javascript/api/excel/excel.worksheetdata#charts)|Возвращает коллекцию диаграмм, имеющихся на листе. Только для чтения.|
||[id](/javascript/api/excel/excel.worksheetdata#id)|Возвращает значение, однозначно идентифицирующее лист в данной книге. Значение идентификатора остается прежним, даже если переименовать или переместить лист. Только для чтения.|
||[name](/javascript/api/excel/excel.worksheetdata#name)|Отображаемое имя листа.|
||[position](/javascript/api/excel/excel.worksheetdata#position)|Положение листа (начиная с нуля) в книге.|
||[Table](/javascript/api/excel/excel.worksheetdata#tables)|Коллекция таблиц, имеющихся на листе. Только для чтения.|
||[доступности](/javascript/api/excel/excel.worksheetdata#visibility)|Видимость листа.|
|[Воркшитлоадоптионс](/javascript/api/excel/excel.worksheetloadoptions)|[$all](/javascript/api/excel/excel.worksheetloadoptions#$all)||
||[темп](/javascript/api/excel/excel.worksheetloadoptions#charts)|Возвращает коллекцию диаграмм, имеющихся на листе.|
||[id](/javascript/api/excel/excel.worksheetloadoptions#id)|Возвращает значение, однозначно идентифицирующее лист в данной книге. Значение идентификатора остается прежним, даже если переименовать или переместить лист. Только для чтения.|
||[name](/javascript/api/excel/excel.worksheetloadoptions#name)|Отображаемое имя листа.|
||[position](/javascript/api/excel/excel.worksheetloadoptions#position)|Положение листа (начиная с нуля) в книге.|
||[Table](/javascript/api/excel/excel.worksheetloadoptions#tables)|Коллекция таблиц, имеющихся на листе.|
||[доступности](/javascript/api/excel/excel.worksheetloadoptions#visibility)|Видимость листа.|
|[Воркшитупдатедата](/javascript/api/excel/excel.worksheetupdatedata)|[name](/javascript/api/excel/excel.worksheetupdatedata#name)|Отображаемое имя листа.|
||[position](/javascript/api/excel/excel.worksheetupdatedata#position)|Положение листа (начиная с нуля) в книге.|
||[visibility](/javascript/api/excel/excel.worksheetupdatedata#visibility)|Видимость листа.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
