---
title: Набор обязательных элементов API JavaScript для Excel 1,6
description: Сведения о наборе требований ExcelApi 1,6
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ea54b825c3d5eff71a3355c3d313feb355d3c892
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430844"
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

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,6. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,6 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,6](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[Суспендапикалкулатионунтилнекстсинк ()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Приостанавливает вычисление до вызова следующего "context.sync()". После этого за пересчет книги и распространение всех зависимостей несет ответственность разработчик.|
|[целлвалуекондитионалформат](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Представляет объект Rule в этом условном форматировании.|
|[колорскалекондитионалформат](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Критерии цветовой шкалы. При использовании цветовой шкалы с двумя координатами средняя точка является необязательной.|
||[сриколорскале](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Если задано значение true, цветовая шкала будет иметь три точки (минимальная, средняя, максимальная), в противном случае будет существовать два (минимум, максимум).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[Formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Оператор условного форматирования текста.|
|[кондитионалколорскалекритериа](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Условие цветовой шкалы "максимальная точка".|
||[точка](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Условие цветовой шкалы "средняя точка", если используется трехцветная цветовая шкала.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Условие цветовой шкалы "минимальная точка".|
|[кондитионалколорскалекритерион](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Цветовое HTML-представление цвета цветовой шкалы. Например, #FF0000 обозначает красный.|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Число, формула или значение NULL (если указан тип LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|Какова должна основываться Условная формула условия.|
|[кондитионалдатабарнегативеформат](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[матчпоситивебордерколор](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет границы, что и положительная.|
||[матчпоситивефиллколор](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Указывает, имеет ли отрицательная гистограмма тот же цвет заливки, что и положительная.|
|[кондитионалдатабарпоситивеформат](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[градиентфилл](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Логическое значение, которое указывает, имеет ли гистограмма градиент.|
|[кондитионалдатабарруле](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Формула, с помощью которой при необходимости оценивается правило гистограммы.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Тип правила для гистограмма.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Удаляет это условное форматирование.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Возврат диапазона, к которому применено условное форматирование. Выдает ошибку, если условное форматирование применяется к нескольким диапазонам. Только для чтения.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Возвращает диапазон, к которому применяется формат кондитонал, или пустой объект, если условное форматирование применяется к нескольким диапазонам. Только для чтения.|
||[важную](/javascript/api/excel/excel.conditionalformat#priority)|Приоритет (или индекс) в коллекции условных форматов, в которой в настоящее время существует данное условное форматирование. Изменение также|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[целлвалуеорнуллобжект](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Возвращает свойства условного форматирования значения ячейки, если текущим условным форматированием является тип CellValue.|
||[Справа](/javascript/api/excel/excel.conditionalformat#colorscale)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы. Только для чтения.|
||[колорскалеорнуллобжект](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Возвращает свойства условного форматирования цветовой шкалы, если текущим условным форматированием является тип цветовой шкалы. Только для чтения.|
||[собственный](/javascript/api/excel/excel.conditionalformat#custom)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип. Только для чтения.|
||[кустоморнуллобжект](/javascript/api/excel/excel.conditionalformat#customornullobject)|Возвращает свойства настраиваемого условного форматирования, если текущим условным форматированием является настраиваемый тип. Только для чтения.|
||[Гистограмма](/javascript/api/excel/excel.conditionalformat#databar)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных. Только для чтения.|
||[датабарорнуллобжект](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Возвращает свойства гистограммы, если текущим условным форматированием является панель данных. Только для чтения.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков. Только для чтения.|
||[иконсеторнуллобжект](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Возвращает свойства условного форматирования набора значков, если текущим условным форматированием является тип набора значков. Только для чтения.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|Приоритет условного форматирования в пределах текущего класса ConditionalFormatCollection. Только для чтения.|
||[набора](/javascript/api/excel/excel.conditionalformat#preset)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[пресеторнуллобжект](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Возвращает условное форматирование предварительно установленных условий. Дополнительные сведения см. в статье Excel. Пресеткритериакондитионалформат.|
||[тексткомпарисон](/javascript/api/excel/excel.conditionalformat#textcomparison)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[тексткомпарисонорнуллобжект](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Возвращает определенные свойства условного форматирования текста, если текущим условным форматированием является текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[топботтоморнуллобжект](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Возвращает верхнее и нижнее свойства условного форматирования, если текущее условное форматирование имеет тип TopBottom.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Тип условного форматирования. В каждый момент времени можно задать только один из них. Только для чтения.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
|[кондитионалформатколлектион](/javascript/api/excel/excel.conditionalformatcollection)|[Добавить (тип: Excel. Кондитионалформаттипе)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Добавляет новое условное форматирование в коллекцию по первому или верхнему приоритету.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Полное удаление условного форматирование в указанном диапазоне.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Возвращает число условных форматов в книге. Только для чтения.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Возвращает условное форматирование для указанного идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Возвращает условное форматирование по индексу.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Формула, с помощью которой при необходимости оценивается правило условного форматирования на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования в формате R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[кустомикон](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Специальный значок для текущего условия, если он отличается от набора значков по умолчанию, в противном случае возвращается значение NULL.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Число или формула в зависимости от типа.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan или Греатерсанорекуал для каждого типа правила для условного форматирования значка.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|На чем должна основываться условная формула значка.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[текущего](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Критерий условного форматирования.|
|[кондитионалранжебордер](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[сидеиндекс](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Кондитионалранжебордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
|[кондитионалранжебордерколлектион](/javascript/api/excel/excel.conditionalrangebordercollection)|[GetItem (index: Excel. Кондитионалранжебордериндекс)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Возвращает объект границы по его индексу.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Получает нижнюю границу. Только для чтения.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Количество объектов границы в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Получает левую границу. Только для чтения.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Получает правую границу. Только для чтения.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Получает верхнюю границу. Только для чтения.|
|[кондитионалранжефилл](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Удаляет заливку.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML-код, представляющий цвет заливки в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
|[кондитионалранжефонт](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Указывает, является ли шрифт полужирным.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Удаляет форматирование шрифтов.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Указывает, применяется ли курсив.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Указывает, зачеркнут ли шрифт.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Кондитионалранжефонтундерлинестиле.|
|[кондитионалранжеформат](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Представляет код числового формата Excel для заданного диапазона. Очищается, если передается значение null.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Коллекция объектов Border, которые применяются к общему диапазону условного форматирования. Только для чтения.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Возвращает объект Fill, определенный в общем диапазоне условного форматирования. Только для чтения.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Возвращает объект Font, определенный в общем диапазоне условного форматирования. Только для чтения.|
|[кондитионалтексткомпарисонруле](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Оператор условного форматирования текста.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Текстовое значение условного форматирования.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|От 1 до 1000 для числовых рейтингов или от 1 до 100 для процентных рейтингов.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Форматирование значений на основе верхнего или нижнего ранга.|
|[кустомкондитионалформат](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.customconditionalformat#rule)|Представляет объект Rule в этом условном форматировании. Только для чтения.|
|[датабаркондитионалформат](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|HTML-код, представляющий цвет линии оси в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[аксисформат](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Представление определения оси для панели данных Excel.|
||[бардиректион](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Представляет направление, на котором должен основываться рисунок на панели данных.|
||[ловербаундруле](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[негативеформат](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Отображение всех значений слева от оси в панели данных Excel. Только для чтения.|
||[поситивеформат](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Представление всех значений справа от оси в панели данных Excel. Только для чтения.|
||[шовдатабаронли](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Значение true скрывает значения ячеек, где применяется гистограмма.|
||[уппербаундруле](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[иконсеткондитионалформат](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Массив критериев и IconSets для правил и потенциальных настраиваемых значков для условных значков. Обратите внимание, что для первого критерия можно изменить только настраиваемый значок, в то время как тип, формула и оператор будут игнорироваться при установке.|
||[реверсеиконордер](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Если этот параметр имеет значение true, отменяет порядок значков для набора значков. Обратите внимание, что этот параметр невозможно задать при использовании настраиваемых значков.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Значение true скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Если этот параметр установлен, отображается параметр "набор значков" для условного форматирования.|
|[пресеткритериакондитионалформат](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства.|
||[правила](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Правило условного форматирования.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Вычисляет диапазон ячеек на листе.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Коллекция объектов Кондитионалформатс, пересекающих диапазон. Только для чтения.|
|[тексткондитионалформат](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.textconditionalformat#rule)|Правило условного форматирования.|
|[топботтомкондитионалформат](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Возвращает объект Format, который инкапсулирует шрифты условного форматирования, заливки, границы и другие свойства. Только для чтения.|
||[правила](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Критерии условного форматирования Top/Bottom.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Calculate (markAllDirty: Boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Вычисляет все ячейки на листе.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
