---
title: Набор обязательных элементов API JavaScript для Excel 1,6
description: Сведения о наборе требований ExcelApi 1,6
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9e1a3375d19d8c1cb0fbddac50fabf826b96d7cc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771976"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Новые возможности API JavaScript для Excel 1.6

## <a name="conditional-formatting"></a>Условное форматирование

Добавлена возможность условного форматирования диапазона. Допускаются следующие типы условного форматирования:

* Цветовая шкала
* Гистограмма
* Набор значков
* Настраиваемый

Дополнительно:

* Возврат диапазона, к которому применено условное форматирование.
* Удаление условного форматирования.
* Возможность использования приоритетов и оператора stopifTrue.
* Получение полной коллекции условного форматирования для определенного диапазона.
* Полное удаление условного форматирование в указанном диапазоне.

## <a name="api-list"></a>Список API

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[Суспендапикалкулатионунтилнекстсинк ()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Приостанавливает вычисление до вызова следующего "context.sync()". После этого за пересчет книги и распространение всех зависимостей несет ответственность разработчик.|
|[Целлвалуекондитионалформат](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Представляет объект Rule в этом условном форматировании.|
||[Set (Properties: Excel. Целлвалуекондитионалформат)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Целлвалуекондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Целлвалуекондитионалформатдата](/javascript/api/excel/excel.cellvalueconditionalformatdata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatdata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.cellvalueconditionalformatdata#rule)|Представляет объект Rule в этом условном форматировании.|
|[Целлвалуекондитионалформатлоадоптионс](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#rule)|Представляет объект Rule в этом условном форматировании.|
|[Целлвалуекондитионалформатупдатедата](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#rule)|Представляет объект Rule в этом условном форматировании.|
|[Колорскалекондитионалформат](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Критерии цветовой шкалы. При использовании цветовой шкалы с двумя координатами средняя точка является необязательной.|
||[Сриколорскале](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Если задано значение true, цветовая шкала будет иметь три точки (минимальная, средняя, максимальная), в противном случае будет существовать два (минимум, максимум).|
||[Set (Properties: Excel. Колорскалекондитионалформат)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Колорскалекондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Колорскалекондитионалформатдата](/javascript/api/excel/excel.colorscaleconditionalformatdata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatdata#criteria)|Критерии цветовой шкалы. При использовании цветовой шкалы с двумя координатами средняя точка является необязательной.|
||[Сриколорскале](/javascript/api/excel/excel.colorscaleconditionalformatdata#threecolorscale)|Если задано значение true, цветовая шкала будет иметь три точки (минимальная, средняя, максимальная), в противном случае будет существовать два (минимум, максимум).|
|[Колорскалекондитионалформатлоадоптионс](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#criteria)|Критерии цветовой шкалы. При использовании цветовой шкалы с двумя координатами средняя точка является необязательной.|
||[Сриколорскале](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#threecolorscale)|Если задано значение true, цветовая шкала будет иметь три точки (минимальная, средняя, максимальная), в противном случае будет существовать два (минимум, максимум).|
|[Колорскалекондитионалформатупдатедата](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata#criteria)|Критерии цветовой шкалы. При использовании цветовой шкалы с двумя координатами средняя точка является необязательной.|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[Formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[or](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Оператор условного форматирования текста.|
|[Кондитионалколорскалекритериа](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Условие цветовой шкалы "максимальная точка".|
||[точка](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Условие цветовой шкалы "средняя точка", если используется трехцветная цветовая шкала.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Условие цветовой шкалы "минимальная точка".|
|[Кондитионалколорскалекритерион](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Цветовое HTML-представление цвета цветовой шкалы. Например, #FF0000 обозначает красный.|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Число, формула или значение NULL (если указан тип LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|Какова должна основываться Условная формула условия.|
|[Кондитионалдатабарнегативеформат](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Матчпоситивебордерколор](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет границы, что и положительная.|
||[Матчпоситивефиллколор](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет заливки, что и положительная.|
||[Set (Properties: Excel. Кондитионалдатабарнегативеформат)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кондитионалдатабарнегативеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Кондитионалдатабарнегативеформатдата](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Матчпоситивебордерколор](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivebordercolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет границы, что и положительная.|
||[Матчпоситивефиллколор](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivefillcolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет заливки, что и положительная.|
|[Кондитионалдатабарнегативеформатлоадоптионс](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Матчпоситивебордерколор](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivebordercolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет границы, что и положительная.|
||[Матчпоситивефиллколор](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivefillcolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет заливки, что и положительная.|
|[Кондитионалдатабарнегативеформатупдатедата](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Матчпоситивебордерколор](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivebordercolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет границы, что и положительная.|
||[Матчпоситивефиллколор](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivefillcolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет заливки, что и положительная.|
|[Кондитионалдатабарпоситивеформат](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Градиентфилл](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Логическое значение, которое указывает, имеет ли гистограмма градиент.|
||[Set (Properties: Excel. Кондитионалдатабарпоситивеформат)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кондитионалдатабарпоситивеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Кондитионалдатабарпоситивеформатдата](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Градиентфилл](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#gradientfill)|Логическое значение, которое указывает, имеет ли гистограмма градиент.|
|[Кондитионалдатабарпоситивеформатлоадоптионс](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Градиентфилл](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#gradientfill)|Логическое значение, которое указывает, имеет ли гистограмма градиент.|
|[Кондитионалдатабарпоситивеформатупдатедата](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Градиентфилл](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#gradientfill)|Логическое значение, которое указывает, имеет ли гистограмма градиент.|
|[Кондитионалдатабарруле](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Формула, с помощью которой при необходимости оценивается правило гистограммы.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Тип правила для гистограмма.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Удаляет это условное форматирование.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Возврат диапазона, к которому применено условное форматирование. Выдает ошибку, если условное форматирование применяется к нескольким диапазонам. Только для чтения.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Возвращает диапазон, к которому применяется формат кондитонал, или пустой объект, если условное форматирование применяется к нескольким диапазонам. Только для чтения.|
||[важную](/javascript/api/excel/excel.conditionalformat#priority)|Приоритет (или индекс) в коллекции условных форматов, в которой в настоящее время существует данное условное форматирование. Изменение также|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Целлвалуеорнуллобжект](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Справа](/javascript/api/excel/excel.conditionalformat#colorscale)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы. Только для чтения.|
||[Колорскалеорнуллобжект](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы. Только для чтения.|
||[собственный](/javascript/api/excel/excel.conditionalformat#custom)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип. Только для чтения.|
||[Кустоморнуллобжект](/javascript/api/excel/excel.conditionalformat#customornullobject)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип. Только для чтения.|
||[Гистограмма](/javascript/api/excel/excel.conditionalformat#databar)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных. Только для чтения.|
||[Датабарорнуллобжект](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных. Только для чтения.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков. Только для чтения.|
||[Иконсеторнуллобжект](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков. Только для чтения.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|Приоритет условного форматирования в пределах текущего класса ConditionalFormatCollection. Только для чтения.|
||[набора](/javascript/api/excel/excel.conditionalformat#preset)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[Пресеторнуллобжект](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[Тексткомпарисон](/javascript/api/excel/excel.conditionalformat#textcomparison)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[Тексткомпарисонорнуллобжект](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[Топботтоморнуллобжект](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Тип условного форматирования. В каждый момент времени можно задать только один из них. Только для чтения.|
||[Set (Properties: Excel. ConditionalFormat)](/javascript/api/excel/excel.conditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.conditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
|[Кондитионалформатколлектион](/javascript/api/excel/excel.conditionalformatcollection)|[Add (Type: "Custom" \| " \| Гистограмма" "" Цветовая \| шкала " \| ", " \| TopBottom" " \| " пресеткритериа " \| " контаинстекст "" CellValue ")](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Добавляет новое условное форматирование в коллекцию по первому или верхнему приоритету.|
||[Добавить (тип: Excel. Кондитионалформаттипе)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Добавляет новое условное форматирование в коллекцию по первому или верхнему приоритету.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Полное удаление условного форматирование в указанном диапазоне.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Возвращает число условных форматов в книге. Только для чтения.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Возвращает условное форматирование для указанного идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Возвращает условное форматирование по индексу.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Кондитионалформатколлектионлоадоптионс](/javascript/api/excel/excel.conditionalformatcollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalue)|Для каждого элемента в коллекции: Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Целлвалуеорнуллобжект](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalueornullobject)|Для каждого элемента в коллекции: Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Справа](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscale)|Для каждого элемента в коллекции: Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы.|
||[Колорскалеорнуллобжект](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscaleornullobject)|Для каждого элемента в коллекции: Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы.|
||[собственный](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#custom)|Для каждого элемента в коллекции: Возвращает свойства настраиваемого условного форматирования, если текущее условное форматирование является настраиваемым типом.|
||[Кустоморнуллобжект](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#customornullobject)|Для каждого элемента в коллекции: Возвращает свойства настраиваемого условного форматирования, если текущее условное форматирование является настраиваемым типом.|
||[Гистограмма](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databar)|Для каждого элемента в коллекции: Возвращает свойства панели данных, если текущим условным форматированием является панель данных.|
||[Датабарорнуллобжект](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databarornullobject)|Для каждого элемента в коллекции: Возвращает свойства панели данных, если текущим условным форматированием является панель данных.|
||[iconSet](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconset)|Для каждого элемента в коллекции: Возвращает свойства условного форматирования набора значков, если текущее условное форматирование является типом набора значков.|
||[Иконсеторнуллобжект](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconsetornullobject)|Для каждого элемента в коллекции: Возвращает свойства условного форматирования набора значков, если текущее условное форматирование является типом набора значков.|
||[id](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#id)|Для каждого элемента в коллекции: приоритет условного форматирования в текущем Кондитионалформатколлектион. Только для чтения.|
||[набора](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#preset)|Для каждого элемента в коллекции: возвращает условное форматирование с предварительно заданными. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[Пресеторнуллобжект](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#presetornullobject)|Для каждого элемента в коллекции: возвращает условное форматирование с предварительно заданными. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[важную](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#priority)|Для каждого элемента в коллекции: приоритет (или индекс) в коллекции условных форматов, в которой в настоящее время существует данное условное форматирование. Изменение также|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#stopiftrue)|Для каждого элемента в коллекции: если условия этого условного форматирования выполнены, то форматы с низким приоритетом не будут применены к этой ячейке.|
||[Тексткомпарисон](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparison)|Для каждого элемента в коллекции: Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[Тексткомпарисонорнуллобжект](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparisonornullobject)|Для каждого элемента в коллекции: Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottom)|Для каждого элемента в коллекции: Возвращает свойства условного форматирования Top/Bottom, если текущее условное форматирование имеет тип TopBottom.|
||[Топботтоморнуллобжект](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottomornullobject)|Для каждого элемента в коллекции: Возвращает свойства условного форматирования Top/Bottom, если текущее условное форматирование имеет тип TopBottom.|
||[type](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#type)|Для каждого элемента в коллекции: тип условного форматирования. В каждый момент времени можно задать только один из них. Только для чтения.|
|[Кондитионалформатдата](/javascript/api/excel/excel.conditionalformatdata)|[cellValue](/javascript/api/excel/excel.conditionalformatdata#cellvalue)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Целлвалуеорнуллобжект](/javascript/api/excel/excel.conditionalformatdata#cellvalueornullobject)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Справа](/javascript/api/excel/excel.conditionalformatdata#colorscale)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы. Только для чтения.|
||[Колорскалеорнуллобжект](/javascript/api/excel/excel.conditionalformatdata#colorscaleornullobject)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы. Только для чтения.|
||[собственный](/javascript/api/excel/excel.conditionalformatdata#custom)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип. Только для чтения.|
||[Кустоморнуллобжект](/javascript/api/excel/excel.conditionalformatdata#customornullobject)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип. Только для чтения.|
||[Гистограмма](/javascript/api/excel/excel.conditionalformatdata#databar)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных. Только для чтения.|
||[Датабарорнуллобжект](/javascript/api/excel/excel.conditionalformatdata#databarornullobject)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных. Только для чтения.|
||[iconSet](/javascript/api/excel/excel.conditionalformatdata#iconset)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков. Только для чтения.|
||[Иконсеторнуллобжект](/javascript/api/excel/excel.conditionalformatdata#iconsetornullobject)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков. Только для чтения.|
||[id](/javascript/api/excel/excel.conditionalformatdata#id)|Приоритет условного форматирования в пределах текущего класса ConditionalFormatCollection. Только для чтения.|
||[набора](/javascript/api/excel/excel.conditionalformatdata#preset)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[Пресеторнуллобжект](/javascript/api/excel/excel.conditionalformatdata#presetornullobject)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[важную](/javascript/api/excel/excel.conditionalformatdata#priority)|Приоритет (или индекс) в коллекции условных форматов, в которой в настоящее время существует данное условное форматирование. Изменение также|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatdata#stopiftrue)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
||[Тексткомпарисон](/javascript/api/excel/excel.conditionalformatdata#textcomparison)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[Тексткомпарисонорнуллобжект](/javascript/api/excel/excel.conditionalformatdata#textcomparisonornullobject)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformatdata#topbottom)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[Топботтоморнуллобжект](/javascript/api/excel/excel.conditionalformatdata#topbottomornullobject)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[type](/javascript/api/excel/excel.conditionalformatdata#type)|Тип условного форматирования. В каждый момент времени можно задать только один из них. Только для чтения.|
|[Кондитионалформатлоадоптионс](/javascript/api/excel/excel.conditionalformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalue)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Целлвалуеорнуллобжект](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalueornullobject)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Справа](/javascript/api/excel/excel.conditionalformatloadoptions#colorscale)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы.|
||[Колорскалеорнуллобжект](/javascript/api/excel/excel.conditionalformatloadoptions#colorscaleornullobject)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы.|
||[собственный](/javascript/api/excel/excel.conditionalformatloadoptions#custom)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип.|
||[Кустоморнуллобжект](/javascript/api/excel/excel.conditionalformatloadoptions#customornullobject)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип.|
||[Гистограмма](/javascript/api/excel/excel.conditionalformatloadoptions#databar)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных.|
||[Датабарорнуллобжект](/javascript/api/excel/excel.conditionalformatloadoptions#databarornullobject)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных.|
||[iconSet](/javascript/api/excel/excel.conditionalformatloadoptions#iconset)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков.|
||[Иконсеторнуллобжект](/javascript/api/excel/excel.conditionalformatloadoptions#iconsetornullobject)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков.|
||[id](/javascript/api/excel/excel.conditionalformatloadoptions#id)|Приоритет условного форматирования в пределах текущего класса ConditionalFormatCollection. Только для чтения.|
||[набора](/javascript/api/excel/excel.conditionalformatloadoptions#preset)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[Пресеторнуллобжект](/javascript/api/excel/excel.conditionalformatloadoptions#presetornullobject)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[важную](/javascript/api/excel/excel.conditionalformatloadoptions#priority)|Приоритет (или индекс) в коллекции условных форматов, в которой в настоящее время существует данное условное форматирование. Изменение также|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatloadoptions#stopiftrue)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
||[Тексткомпарисон](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparison)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[Тексткомпарисонорнуллобжект](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparisonornullobject)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformatloadoptions#topbottom)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[Топботтоморнуллобжект](/javascript/api/excel/excel.conditionalformatloadoptions#topbottomornullobject)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[type](/javascript/api/excel/excel.conditionalformatloadoptions#type)|Тип условного форматирования. В каждый момент времени можно задать только один из них. Только для чтения.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Формула, с помощью которой при необходимости оценивается правило условного форматирования на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования в формате R1C1.|
||[Set (Properties: Excel. ConditionalFormatRule)](/javascript/api/excel/excel.conditionalformatrule#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кондитионалформатрулеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.conditionalformatrule#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Кондитионалформатруледата](/javascript/api/excel/excel.conditionalformatruledata)|[formula](/javascript/api/excel/excel.conditionalformatruledata#formula)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruledata#formulalocal)|Формула, с помощью которой при необходимости оценивается правило условного форматирования на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruledata#formular1c1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования в формате R1C1.|
|[Кондитионалформатрулелоадоптионс](/javascript/api/excel/excel.conditionalformatruleloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatruleloadoptions#$all)||
||[formula](/javascript/api/excel/excel.conditionalformatruleloadoptions#formula)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleloadoptions#formulalocal)|Формула, с помощью которой при необходимости оценивается правило условного форматирования на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleloadoptions#formular1c1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования в формате R1C1.|
|[Кондитионалформатрулеупдатедата](/javascript/api/excel/excel.conditionalformatruleupdatedata)|[formula](/javascript/api/excel/excel.conditionalformatruleupdatedata#formula)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleupdatedata#formulalocal)|Формула, с помощью которой при необходимости оценивается правило условного форматирования на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleupdatedata#formular1c1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования в формате R1C1.|
|[Кондитионалформатупдатедата](/javascript/api/excel/excel.conditionalformatupdatedata)|[cellValue](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalue)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Целлвалуеорнуллобжект](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalueornullobject)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Справа](/javascript/api/excel/excel.conditionalformatupdatedata#colorscale)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы.|
||[Колорскалеорнуллобжект](/javascript/api/excel/excel.conditionalformatupdatedata#colorscaleornullobject)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы.|
||[собственный](/javascript/api/excel/excel.conditionalformatupdatedata#custom)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип.|
||[Кустоморнуллобжект](/javascript/api/excel/excel.conditionalformatupdatedata#customornullobject)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип.|
||[Гистограмма](/javascript/api/excel/excel.conditionalformatupdatedata#databar)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных.|
||[Датабарорнуллобжект](/javascript/api/excel/excel.conditionalformatupdatedata#databarornullobject)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных.|
||[iconSet](/javascript/api/excel/excel.conditionalformatupdatedata#iconset)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков.|
||[Иконсеторнуллобжект](/javascript/api/excel/excel.conditionalformatupdatedata#iconsetornullobject)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков.|
||[набора](/javascript/api/excel/excel.conditionalformatupdatedata#preset)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[Пресеторнуллобжект](/javascript/api/excel/excel.conditionalformatupdatedata#presetornullobject)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[важную](/javascript/api/excel/excel.conditionalformatupdatedata#priority)|Приоритет (или индекс) в коллекции условных форматов, в которой в настоящее время существует данное условное форматирование. Изменение также|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatupdatedata#stopiftrue)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
||[Тексткомпарисон](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparison)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[Тексткомпарисонорнуллобжект](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparisonornullobject)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformatupdatedata#topbottom)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[Топботтоморнуллобжект](/javascript/api/excel/excel.conditionalformatupdatedata#topbottomornullobject)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[Кустомикон](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Специальный значок для текущего условия, если он отличается от набора значков по умолчанию, в противном случае возвращается значение NULL.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Число или формула в зависимости от типа.|
||[or](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan или Греатерсанорекуал для каждого типа правила для условного форматирования значка.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|На чем должна основываться условная формула значка.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[текущего](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Критерий условного форматирования.|
|[Кондитионалранжебордер](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Сидеиндекс](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Кондитионалранжебордериндекс. Только для чтения.|
||[Set (Properties: Excel. Кондитионалранжебордер)](/javascript/api/excel/excel.conditionalrangeborder#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кондитионалранжебордерупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.conditionalrangeborder#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
|[Кондитионалранжебордерколлектион](/javascript/api/excel/excel.conditionalrangebordercollection)|[GetItem (index: "Еджетоп" \| "еджеботтом" \| "еджелефт" \| "еджеригхт")](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[GetItem (index: Excel. Кондитионалранжебордериндекс)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Возвращает объект границы по его индексу.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Получает нижнюю границу. Только для чтения.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Количество объектов границы в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Получает левую границу. Только для чтения.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Получает правую границу. Только для чтения.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Получает верхнюю границу. Только для чтения.|
|[Кондитионалранжебордерколлектионлоадоптионс](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#color)|Для каждого элемента в коллекции: HTML-код цвета, представляющий цвет линии границы, формы #RRGGBB (например, "FFA500") или в виде именованного цвета HTML (например, "Апельсин").|
||[Сидеиндекс](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#sideindex)|Для каждого элемента в коллекции: значение константы, которое указывает на конкретную сторону границы. Дополнительные сведения см. в статье Excel. Кондитионалранжебордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#style)|Для каждого элемента в коллекции: одна из констант стиля линии, определяющая стиль линии для границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
|[Кондитионалранжебордерколлектионупдатедата](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#bottom)|Получает нижнюю границу.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#left)|Получает левую границу.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#right)|Получает правую границу.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#top)|Получает верхнюю границу.|
|[Кондитионалранжебордердата](/javascript/api/excel/excel.conditionalrangeborderdata)|[color](/javascript/api/excel/excel.conditionalrangeborderdata#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Сидеиндекс](/javascript/api/excel/excel.conditionalrangeborderdata#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Кондитионалранжебордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.conditionalrangeborderdata#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
|[Кондитионалранжебордерлоадоптионс](/javascript/api/excel/excel.conditionalrangeborderloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangeborderloadoptions#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Сидеиндекс](/javascript/api/excel/excel.conditionalrangeborderloadoptions#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Кондитионалранжебордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.conditionalrangeborderloadoptions#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
|[Кондитионалранжебордерупдатедата](/javascript/api/excel/excel.conditionalrangeborderupdatedata)|[color](/javascript/api/excel/excel.conditionalrangeborderupdatedata#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[style](/javascript/api/excel/excel.conditionalrangeborderupdatedata#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
|[Кондитионалранжефилл](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Удаляет заливку.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Set (Properties: Excel. Кондитионалранжефилл)](/javascript/api/excel/excel.conditionalrangefill#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кондитионалранжефиллупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.conditionalrangefill#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Кондитионалранжефиллдата](/javascript/api/excel/excel.conditionalrangefilldata)|[color](/javascript/api/excel/excel.conditionalrangefilldata#color)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
|[Кондитионалранжефилллоадоптионс](/javascript/api/excel/excel.conditionalrangefillloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangefillloadoptions#color)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
|[Кондитионалранжефиллупдатедата](/javascript/api/excel/excel.conditionalrangefillupdatedata)|[color](/javascript/api/excel/excel.conditionalrangefillupdatedata#color)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
|[Кондитионалранжефонт](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Указывает, является ли шрифт полужирным.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Удаляет форматирование шрифтов.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Указывает, применяется ли курсив.|
||[Set (Properties: Excel. Кондитионалранжефонт)](/javascript/api/excel/excel.conditionalrangefont#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кондитионалранжефонтупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.conditionalrangefont#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Указывает, зачеркнут ли шрифт.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Кондитионалранжефонтундерлинестиле.|
|[Кондитионалранжефонтдата](/javascript/api/excel/excel.conditionalrangefontdata)|[bold](/javascript/api/excel/excel.conditionalrangefontdata#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.conditionalrangefontdata#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.conditionalrangefontdata#italic)|Указывает, применяется ли курсив.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontdata#strikethrough)|Указывает, зачеркнут ли шрифт.|
||[underline](/javascript/api/excel/excel.conditionalrangefontdata#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Кондитионалранжефонтундерлинестиле.|
|[Кондитионалранжефонтлоадоптионс](/javascript/api/excel/excel.conditionalrangefontloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.conditionalrangefontloadoptions#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.conditionalrangefontloadoptions#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.conditionalrangefontloadoptions#italic)|Указывает, применяется ли курсив.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontloadoptions#strikethrough)|Указывает, зачеркнут ли шрифт.|
||[underline](/javascript/api/excel/excel.conditionalrangefontloadoptions#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Кондитионалранжефонтундерлинестиле.|
|[Кондитионалранжефонтупдатедата](/javascript/api/excel/excel.conditionalrangefontupdatedata)|[bold](/javascript/api/excel/excel.conditionalrangefontupdatedata#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.conditionalrangefontupdatedata#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.conditionalrangefontupdatedata#italic)|Указывает, применяется ли курсив.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontupdatedata#strikethrough)|Указывает, зачеркнут ли шрифт.|
||[underline](/javascript/api/excel/excel.conditionalrangefontupdatedata#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Кондитионалранжефонтундерлинестиле.|
|[Кондитионалранжеформат](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Представляет код числового формата Excel для заданного диапазона. Очищается, если передается значение null.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Коллекция объектов Border, которые применяются к общему диапазону условного форматирования. Только для чтения.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Возвращает объект Fill, определенный в общем диапазоне условного форматирования. Только для чтения.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Возвращает объект Font, определенный в общем диапазоне условного форматирования. Только для чтения.|
||[Set (Properties: Excel. Кондитионалранжеформат)](/javascript/api/excel/excel.conditionalrangeformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кондитионалранжеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.conditionalrangeformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Кондитионалранжеформатдата](/javascript/api/excel/excel.conditionalrangeformatdata)|[borders](/javascript/api/excel/excel.conditionalrangeformatdata#borders)|Коллекция объектов Border, которые применяются к общему диапазону условного форматирования. Только для чтения.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatdata#fill)|Возвращает объект Fill, определенный в общем диапазоне условного форматирования. Только для чтения.|
||[font](/javascript/api/excel/excel.conditionalrangeformatdata#font)|Возвращает объект Font, определенный в общем диапазоне условного форматирования. Только для чтения.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatdata#numberformat)|Представляет код числового формата Excel для заданного диапазона. Очищается, если передается значение null.|
|[Кондитионалранжеформатлоадоптионс](/javascript/api/excel/excel.conditionalrangeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeformatloadoptions#$all)||
||[borders](/javascript/api/excel/excel.conditionalrangeformatloadoptions#borders)|Коллекция объектов Border, которые применяются к общему диапазону условного форматирования.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatloadoptions#fill)|Возвращает объект Fill, определенный в общем диапазоне условного форматирования.|
||[font](/javascript/api/excel/excel.conditionalrangeformatloadoptions#font)|Возвращает объект Font, определенный в общем диапазоне условного форматирования.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatloadoptions#numberformat)|Представляет код числового формата Excel для заданного диапазона. Очищается, если передается значение null.|
|[Кондитионалранжеформатупдатедата](/javascript/api/excel/excel.conditionalrangeformatupdatedata)|[borders](/javascript/api/excel/excel.conditionalrangeformatupdatedata#borders)|Коллекция объектов Border, которые применяются к общему диапазону условного форматирования.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatupdatedata#fill)|Возвращает объект Fill, определенный в общем диапазоне условного форматирования.|
||[font](/javascript/api/excel/excel.conditionalrangeformatupdatedata#font)|Возвращает объект Font, определенный в общем диапазоне условного форматирования.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatupdatedata#numberformat)|Представляет код числового формата Excel для заданного диапазона. Очищается, если передается значение null.|
|[Кондитионалтексткомпарисонруле](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[or](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Оператор условного форматирования текста.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Текстовое значение условного форматирования.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|От 1 до 1000 для числовых рейтингов или от 1 до 100 для процентных рейтингов.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Форматирование значений на основе верхнего или нижнего ранга.|
|[Кустомкондитионалформат](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.customconditionalformat#rule)|Представляет объект Rule в этом условном форматировании. Только для чтения.|
||[Set (Properties: Excel. Кустомкондитионалформат)](/javascript/api/excel/excel.customconditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Кустомкондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.customconditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Кустомкондитионалформатдата](/javascript/api/excel/excel.customconditionalformatdata)|[format](/javascript/api/excel/excel.customconditionalformatdata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.customconditionalformatdata#rule)|Представляет объект Rule в этом условном форматировании. Только для чтения.|
|[Кустомкондитионалформатлоадоптионс](/javascript/api/excel/excel.customconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.customconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.customconditionalformatloadoptions#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.customconditionalformatloadoptions#rule)|Представляет объект Rule в этом условном форматировании.|
|[Кустомкондитионалформатупдатедата](/javascript/api/excel/excel.customconditionalformatupdatedata)|[format](/javascript/api/excel/excel.customconditionalformatupdatedata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.customconditionalformatupdatedata#rule)|Представляет объект Rule в этом условном форматировании.|
|[Датабаркондитионалформат](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|HTML-код, представляющий цвет линии оси в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Аксисформат](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Представление определения оси для панели данных Excel.|
||[Бардиректион](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Представляет направление, на котором должен основываться рисунок на панели данных.|
||[Ловербаундруле](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[Негативеформат](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Отображение всех значений слева от оси в панели данных Excel. Только для чтения.|
||[Поситивеформат](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Представление всех значений справа от оси в панели данных Excel. Только для чтения.|
||[Set (Properties: Excel. Датабаркондитионалформат)](/javascript/api/excel/excel.databarconditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Датабаркондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.databarconditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[Шовдатабаронли](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Значение true скрывает значения ячеек, где применяется гистограмма.|
||[Уппербаундруле](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[Датабаркондитионалформатдата](/javascript/api/excel/excel.databarconditionalformatdata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatdata#axiscolor)|HTML-код, представляющий цвет линии оси в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Аксисформат](/javascript/api/excel/excel.databarconditionalformatdata#axisformat)|Представление определения оси для панели данных Excel.|
||[Бардиректион](/javascript/api/excel/excel.databarconditionalformatdata#bardirection)|Представляет направление, на котором должен основываться рисунок на панели данных.|
||[Ловербаундруле](/javascript/api/excel/excel.databarconditionalformatdata#lowerboundrule)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[Негативеформат](/javascript/api/excel/excel.databarconditionalformatdata#negativeformat)|Отображение всех значений слева от оси в панели данных Excel. Только для чтения.|
||[Поситивеформат](/javascript/api/excel/excel.databarconditionalformatdata#positiveformat)|Представление всех значений справа от оси в панели данных Excel. Только для чтения.|
||[Шовдатабаронли](/javascript/api/excel/excel.databarconditionalformatdata#showdatabaronly)|Значение true скрывает значения ячеек, где применяется гистограмма.|
||[Уппербаундруле](/javascript/api/excel/excel.databarconditionalformatdata#upperboundrule)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[Датабаркондитионалформатлоадоптионс](/javascript/api/excel/excel.databarconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.databarconditionalformatloadoptions#$all)||
||[axisColor](/javascript/api/excel/excel.databarconditionalformatloadoptions#axiscolor)|HTML-код, представляющий цвет линии оси в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Аксисформат](/javascript/api/excel/excel.databarconditionalformatloadoptions#axisformat)|Представление определения оси для панели данных Excel.|
||[Бардиректион](/javascript/api/excel/excel.databarconditionalformatloadoptions#bardirection)|Представляет направление, на котором должен основываться рисунок на панели данных.|
||[Ловербаундруле](/javascript/api/excel/excel.databarconditionalformatloadoptions#lowerboundrule)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[Негативеформат](/javascript/api/excel/excel.databarconditionalformatloadoptions#negativeformat)|Отображение всех значений слева от оси в панели данных Excel.|
||[Поситивеформат](/javascript/api/excel/excel.databarconditionalformatloadoptions#positiveformat)|Представление всех значений справа от оси в панели данных Excel.|
||[Шовдатабаронли](/javascript/api/excel/excel.databarconditionalformatloadoptions#showdatabaronly)|Значение true скрывает значения ячеек, где применяется гистограмма.|
||[Уппербаундруле](/javascript/api/excel/excel.databarconditionalformatloadoptions#upperboundrule)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[Датабаркондитионалформатупдатедата](/javascript/api/excel/excel.databarconditionalformatupdatedata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatupdatedata#axiscolor)|HTML-код, представляющий цвет линии оси в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Аксисформат](/javascript/api/excel/excel.databarconditionalformatupdatedata#axisformat)|Представление определения оси для панели данных Excel.|
||[Бардиректион](/javascript/api/excel/excel.databarconditionalformatupdatedata#bardirection)|Представляет направление, на котором должен основываться рисунок на панели данных.|
||[Ловербаундруле](/javascript/api/excel/excel.databarconditionalformatupdatedata#lowerboundrule)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[Негативеформат](/javascript/api/excel/excel.databarconditionalformatupdatedata#negativeformat)|Отображение всех значений слева от оси в панели данных Excel.|
||[Поситивеформат](/javascript/api/excel/excel.databarconditionalformatupdatedata#positiveformat)|Представление всех значений справа от оси в панели данных Excel.|
||[Шовдатабаронли](/javascript/api/excel/excel.databarconditionalformatupdatedata#showdatabaronly)|Значение true скрывает значения ячеек, где применяется гистограмма.|
||[Уппербаундруле](/javascript/api/excel/excel.databarconditionalformatupdatedata#upperboundrule)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[Иконсеткондитионалформат](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Массив критериев и IconSets для правил и потенциальных настраиваемых значков для условных значков. Обратите внимание, что для первого критерия можно изменить только настраиваемый значок, в то время как тип, формула и оператор будут игнорироваться при установке.|
||[Реверсеиконордер](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Если этот параметр имеет значение true, отменяет порядок значков для набора значков. Обратите внимание, что этот параметр невозможно задать при использовании настраиваемых значков.|
||[Set (Properties: Excel. Иконсеткондитионалформат)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Иконсеткондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Значение true скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Если этот параметр установлен, отображается параметр "набор значков" для условного форматирования.|
|[Иконсеткондитионалформатдата](/javascript/api/excel/excel.iconsetconditionalformatdata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatdata#criteria)|Массив критериев и IconSets для правил и потенциальных настраиваемых значков для условных значков. Обратите внимание, что для первого критерия можно изменить только настраиваемый значок, в то время как тип, формула и оператор будут игнорироваться при установке.|
||[Реверсеиконордер](/javascript/api/excel/excel.iconsetconditionalformatdata#reverseiconorder)|Если этот параметр имеет значение true, отменяет порядок значков для набора значков. Обратите внимание, что этот параметр невозможно задать при использовании настраиваемых значков.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatdata#showicononly)|Значение true скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatdata#style)|Если этот параметр установлен, отображается параметр "набор значков" для условного форматирования.|
|[Иконсеткондитионалформатлоадоптионс](/javascript/api/excel/excel.iconsetconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#criteria)|Массив критериев и IconSets для правил и потенциальных настраиваемых значков для условных значков. Обратите внимание, что для первого критерия можно изменить только настраиваемый значок, в то время как тип, формула и оператор будут игнорироваться при установке.|
||[Реверсеиконордер](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#reverseiconorder)|Если этот параметр имеет значение true, отменяет порядок значков для набора значков. Обратите внимание, что этот параметр невозможно задать при использовании настраиваемых значков.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#showicononly)|Значение true скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#style)|Если этот параметр установлен, отображается параметр "набор значков" для условного форматирования.|
|[Иконсеткондитионалформатупдатедата](/javascript/api/excel/excel.iconsetconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#criteria)|Массив критериев и IconSets для правил и потенциальных настраиваемых значков для условных значков. Обратите внимание, что для первого критерия можно изменить только настраиваемый значок, в то время как тип, формула и оператор будут игнорироваться при установке.|
||[Реверсеиконордер](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#reverseiconorder)|Если этот параметр имеет значение true, отменяет порядок значков для набора значков. Обратите внимание, что этот параметр невозможно задать при использовании настраиваемых значков.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#showicononly)|Значение true скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#style)|Если этот параметр установлен, отображается параметр "набор значков" для условного форматирования.|
|[Пресеткритериакондитионалформат](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Правило условного форматирования.|
||[Set (Properties: Excel. Пресеткритериакондитионалформат)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Пресеткритериакондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Пресеткритериакондитионалформатдата](/javascript/api/excel/excel.presetcriteriaconditionalformatdata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#rule)|Правило условного форматирования.|
|[Пресеткритериакондитионалформатлоадоптионс](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#rule)|Правило условного форматирования.|
|[Пресеткритериакондитионалформатупдатедата](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#rule)|Правило условного форматирования.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Вычисляет диапазон ячеек на листе.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Коллекция объектов Кондитионалформатс, пересекающих диапазон. Только для чтения.|
|[Ранжедата](/javascript/api/excel/excel.rangedata)|[conditionalFormats](/javascript/api/excel/excel.rangedata#conditionalformats)|Коллекция объектов Кондитионалформатс, пересекающих диапазон. Только для чтения.|
|[Тексткондитионалформат](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.textconditionalformat#rule)|Правило условного форматирования.|
||[Set (Properties: Excel. Тексткондитионалформат)](/javascript/api/excel/excel.textconditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Тексткондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.textconditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Тексткондитионалформатдата](/javascript/api/excel/excel.textconditionalformatdata)|[format](/javascript/api/excel/excel.textconditionalformatdata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.textconditionalformatdata#rule)|Правило условного форматирования.|
|[Тексткондитионалформатлоадоптионс](/javascript/api/excel/excel.textconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.textconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.textconditionalformatloadoptions#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.textconditionalformatloadoptions#rule)|Правило условного форматирования.|
|[Тексткондитионалформатупдатедата](/javascript/api/excel/excel.textconditionalformatupdatedata)|[format](/javascript/api/excel/excel.textconditionalformatupdatedata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.textconditionalformatupdatedata#rule)|Правило условного форматирования.|
|[Топботтомкондитионалформат](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Критерии условного форматирования Top/Bottom.|
||[Set (Properties: Excel. Топботтомкондитионалформат)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Топботтомкондитионалформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Топботтомкондитионалформатдата](/javascript/api/excel/excel.topbottomconditionalformatdata)|[format](/javascript/api/excel/excel.topbottomconditionalformatdata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.topbottomconditionalformatdata#rule)|Критерии условного форматирования Top/Bottom.|
|[Топботтомкондитионалформатлоадоптионс](/javascript/api/excel/excel.topbottomconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#rule)|Критерии условного форматирования Top/Bottom.|
|[Топботтомкондитионалформатупдатедата](/javascript/api/excel/excel.topbottomconditionalformatupdatedata)|[format](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#rule)|Критерии условного форматирования Top/Bottom.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Calculate (markAllDirty: Boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Вычисляет все ячейки на листе.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
