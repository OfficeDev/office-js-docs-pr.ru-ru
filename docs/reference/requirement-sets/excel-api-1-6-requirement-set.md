---
title: Набор обязательных элементов API JavaScript для Excel 1,6
description: Сведения о наборе требований ExcelApi 1,6
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c38dd942c3002af05f847846145bc89f1cbbccbe
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064909"
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
* Предоставляет приоритет и `stopifTrue` возможности.
* Получение полной коллекции условного форматирования для определенного диапазона.
* Полное удаление условного форматирование в указанном диапазоне.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,6. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,6 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,6](/javascript/api/excel?view=excel-js-1.6)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[Суспендапикалкулатионунтилнекстсинк ()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Приостанавливает вычисление до вызова следующего "context.sync()". После этого за пересчет книги и распространение всех зависимостей несет ответственность разработчик.|
|[Целлвалуекондитионалформат](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Представляет объект Rule в этом условном форматировании.|
|[Колорскалекондитионалформат](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Критерии цветовой шкалы. При использовании цветовой шкалы с двумя координатами средняя точка является необязательной.|
||[Сриколорскале](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Если задано значение true, цветовая шкала будет иметь три точки (минимальная, средняя, максимальная), в противном случае будет существовать два (минимум, максимум).|
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
|[Кондитионалдатабарпоситивеформат](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Градиентфилл](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Логическое значение, которое указывает, имеет ли гистограмма градиент.|
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
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
|[Кондитионалформатколлектион](/javascript/api/excel/excel.conditionalformatcollection)|[Добавить (тип: Excel. Кондитионалформаттипе)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Добавляет новое условное форматирование в коллекцию по первому или верхнему приоритету.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Полное удаление условного форматирование в указанном диапазоне.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Возвращает число условных форматов в книге. Только для чтения.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Возвращает условное форматирование для указанного идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Возвращает условное форматирование по индексу.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Формула, с помощью которой при необходимости оценивается правило условного форматирования на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования в формате R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[Кустомикон](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Специальный значок для текущего условия, если он отличается от набора значков по умолчанию, в противном случае возвращается значение NULL.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Число или формула в зависимости от типа.|
||[or](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan или Греатерсанорекуал для каждого типа правила для условного форматирования значка.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|На чем должна основываться условная формула значка.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[текущего](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Критерий условного форматирования.|
|[Кондитионалранжебордер](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Сидеиндекс](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Кондитионалранжебордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
|[Кондитионалранжебордерколлектион](/javascript/api/excel/excel.conditionalrangebordercollection)|[GetItem (index: Excel. Кондитионалранжебордериндекс)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Возвращает объект границы по его индексу.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Получает нижнюю границу. Только для чтения.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Количество объектов границы в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Получает левую границу. Только для чтения.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Получает правую границу. Только для чтения.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Получает верхнюю границу. Только для чтения.|
|[Кондитионалранжефилл](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Удаляет заливку.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
|[Кондитионалранжефонт](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Указывает, является ли шрифт полужирным.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Удаляет форматирование шрифтов.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Указывает, применяется ли курсив.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Указывает, зачеркнут ли шрифт.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Кондитионалранжефонтундерлинестиле.|
|[Кондитионалранжеформат](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Представляет код числового формата Excel для заданного диапазона. Очищается, если передается значение null.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Коллекция объектов Border, которые применяются к общему диапазону условного форматирования. Только для чтения.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Возвращает объект Fill, определенный в общем диапазоне условного форматирования. Только для чтения.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Возвращает объект Font, определенный в общем диапазоне условного форматирования. Только для чтения.|
|[Кондитионалтексткомпарисонруле](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[or](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Оператор условного форматирования текста.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Текстовое значение условного форматирования.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|От 1 до 1000 для числовых рейтингов или от 1 до 100 для процентных рейтингов.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Форматирование значений на основе верхнего или нижнего ранга.|
|[Кустомкондитионалформат](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.customconditionalformat#rule)|Представляет объект Rule в этом условном форматировании. Только для чтения.|
|[Датабаркондитионалформат](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|HTML-код, представляющий цвет линии оси в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Аксисформат](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Представление определения оси для панели данных Excel.|
||[Бардиректион](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Представляет направление, на котором должен основываться рисунок на панели данных.|
||[Ловербаундруле](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[Негативеформат](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Отображение всех значений слева от оси в панели данных Excel. Только для чтения.|
||[Поситивеформат](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Представление всех значений справа от оси в панели данных Excel. Только для чтения.|
||[Шовдатабаронли](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Значение true скрывает значения ячеек, где применяется гистограмма.|
||[Уппербаундруле](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[Иконсеткондитионалформат](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Массив критериев и IconSets для правил и потенциальных настраиваемых значков для условных значков. Обратите внимание, что для первого критерия можно изменить только настраиваемый значок, в то время как тип, формула и оператор будут игнорироваться при установке.|
||[Реверсеиконордер](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Если этот параметр имеет значение true, отменяет порядок значков для набора значков. Обратите внимание, что этот параметр невозможно задать при использовании настраиваемых значков.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Значение true скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Если этот параметр установлен, отображается параметр "набор значков" для условного форматирования.|
|[Пресеткритериакондитионалформат](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Правило условного форматирования.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Вычисляет диапазон ячеек на листе.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Коллекция объектов Кондитионалформатс, пересекающих диапазон. Только для чтения.|
|[Тексткондитионалформат](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.textconditionalformat#rule)|Правило условного форматирования.|
|[Топботтомкондитионалформат](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Критерии условного форматирования Top/Bottom.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Calculate (markAllDirty: Boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Вычисляет все ячейки на листе.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.6)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
