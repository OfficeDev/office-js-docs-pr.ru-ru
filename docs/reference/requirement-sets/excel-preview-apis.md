---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих Excel API JavaScript.
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d90c5e8bb2c344cb3bb297a3cd793613f017e910ab99df6dfffc456c3f715d20
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092648"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

В следующей таблице приводится краткий сводка [API,](#api-list) а в следующей таблице списка API приводится подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Таблицы данных диаграммы | Управление внешним видом, форматированием и видимостью таблиц данных на диаграммах. | [Диаграмма](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Задачи документа | Превратите комментарии в задачи, назначенные пользователям. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Удостоверения | Управление удостоверениями пользователей, включая имя отображения и адрес электронной почты. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Связанные типы данных | Добавляет поддержку типов данных, подключенных к Excel из внешних источников. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Связанные книги | Управление связями между книгами, включая поддержку обновления и разрыва ссылок на книги. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Стили таблиц | Обеспечивает управление шрифтом, границей, цветом заполнения и другими аспектами стилей таблиц. | [Таблица](/javascript/api/excel/excel.table), [PivotTable](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Запросы | Извлечение атрибутов запроса, таких как имя, дата обновления и количество запросов. | [Запрос](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены Excel API JavaScript, которые в настоящее время находятся в предварительном просмотре. Полный список всех API Excel JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. Excel [API JavaScript.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Представляет направление (например, вверх или влево), которое остальные ячейки будут смещаться при удалении ячейки или ячейки.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Представляет направление (например, вниз или вправо), в которое будут перенесены существующие ячейки при вставке новой ячейки или ячеек.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Получает таблицу данных на диаграмме.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Получает таблицу данных на диаграмме.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Представляет формат таблицы данных диаграммы, которая включает заполняемую таблицу, шрифт и пограничный формат.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Указывает, следует ли отображать горизонтальную границу таблицы данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Указывает, следует ли показывать legendkey таблицы данных.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Указывает, следует ли отображать контурную границу таблицы данных.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Указывает, следует ли отображать вертикальную границу таблицы данных.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Указывает, следует ли показывать таблицу данных диаграммы.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[граница](/javascript/api/excel/excel.chartdatatableformat#border)|Представляет пограничный формат таблицы данных диаграммы, которая включает цвет, стиль строки и вес.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|Представляет атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для текущего объекта.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве ассимилята.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Получает задачу, связанную с этим комментарием.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Получает задачу, связанную с этим комментарием.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Получает примечание из коллекции на основе его идентификатора.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве единственного назначаемой.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Получает задачу, связанную с потоком ответа на этот комментарий.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Получает задачу, связанную с потоком ответа на этот комментарий.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Возвращает ответ на примечание, определенное по идентификатору.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Возвращает условный формат, идентифицированный его ID.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Указывает процент выполнения задачи.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Указывает приоритет задачи.|
||[назначение](/javascript/api/excel/excel.documenttask#assignees)|Возвращает коллекцию назначений задачи.|
||[изменения](/javascript/api/excel/excel.documenttask#changes)|Получает записи изменений задачи.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Получает комментарий, связанный с задачей.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|Получает последнего пользователя, который выполнил задачу.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|Получает дату и время завершения задачи.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|Получает пользователя, создавшего задачу.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|Получает дату и время создания задачи.|
||[id](/javascript/api/excel/excel.documenttask#id)|Получает ID задачи.|
||[setStartAndDueDateTime (startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setStartAndDueDateTime_startDateTime__dueDateTime_)|Изменяет даты начала и срока действия задачи.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startAndDueDateTime)|Получает или задает дату и время, когда должна начаться и должна быть поставлена задача.|
||[заголовок](/javascript/api/excel/excel.documenttask#title)|Указывает название задачи.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Представляет пользователя, назначенного для задачи для типа записи изменений, или пользователя, не назначенного из задачи `assign` для типа записи `unassign` изменений.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedBy)|Представляет пользователя, создавшего или измениввшего задачу.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentId)|Представляет ID того или иного `Comment` изменения `CommentReply` задачи.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createdDateTime)|Представляет дату создания и время записи изменения задачи.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#dueDateTime)|Представляет дату и время задачи в часовом поясе UTC.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID для записи изменения задачи.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentComplete)|Представляет процент выполнения задачи.|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|Представляет приоритет задачи.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startDateTime)|Представляет дату и время начала задачи в часовом поясе UTC.|
||[заголовок](/javascript/api/excel/excel.documenttaskchange#title)|Представляет название задачи.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Представляет тип действия записи изменения задачи.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undoHistoryId)|Представляет `DocumentTaskChange.id` свойство, которое было отменено для типа `undo` записи изменений.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getCount__)|Получает количество записей изменений в коллекции для задачи.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getItemAt_index_)|Получает запись изменения задачи с помощью индекса в коллекции.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getCount__)|Получает количество задач в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItem_key_)|Получает задачу с помощью своего ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getItemAt_index_)|Получает задачу по индексу в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItemOrNullObject_key_)|Получает задачу с помощью своего ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#dueDateTime)|Получает дату и время, когда должна быть поставлена задача.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startDateTime)|Получает дату и время, которые должна начаться задача.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|Получает фигуру с ее именем или ИД.|
|[Удостоверение](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayName)|Представляет отображаемое имя пользователя.|
||[email](/javascript/api/excel/excel.identity#email)|Представляет электронный адрес пользователя.|
||[id](/javascript/api/excel/excel.identity#id)|Представляет уникальный ID пользователя.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add_assignee_)|Добавляет идентификатор пользователя в коллекцию.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear__)|Удаляет все идентификаторы пользователей из коллекции.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getCount__)|Возвращает число элементов в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getItemAt_index_)|Получает удостоверение пользователя документа с помощью индекса в коллекции.|
||[items](/javascript/api/excel/excel.identitycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove_assignee_)|Удаляет удостоверение пользователя из коллекции.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayName)|Представляет отображаемое имя пользователя.|
||[email](/javascript/api/excel/excel.identityentity#email)|Представляет электронный адрес пользователя.|
||[id](/javascript/api/excel/excel.identityentity#id)|Представляет уникальный ID пользователя.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataProvider)|Имя поставщика данных для связанного типа данных.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastRefreshed)|Дата и время локального часового пояса с момента открытия книги при последнем обновлении связанного типа данных.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Имя связанного типа данных.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicRefreshInterval)|Частота в секундах, при которой тип связанных данных обновляется, если `refreshMode` установлено "Периодическое".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshMode)|Механизм получения данных для связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|Уникальный ID связанного типа данных.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|Возвращает массив со всеми режимами обновления, поддерживаемыми типом связанных данных.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|Делает запрос на обновление связанного типа данных.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|Делает запрос на изменение режима обновления для этого связанного типа данных.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|Уникальный ID нового типа связанных данных.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Получает тип события.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|Получает количество связанных типов данных в коллекции.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|Получает связанный тип данных по ID службы.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|Получает связанный тип данных по индексу в коллекции.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|Получает связанный тип данных по ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|Делает запрос на обновление всех связанных типов данных в коллекции.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Делает запрос на разрыв ссылок, указывающих на связанную книгу.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|Исходный URL-адрес, указывающий на связанную книгу.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Делает запрос на обновление данных, извлеченных из связанной книги.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Нарушает все ссылки на связанные книги.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Получает сведения о связанной книге по URL-адресу.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Получает сведения о связанной книге по URL-адресу.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Делает запрос на обновление всех ссылок на книги.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Представляет режим обновления ссылок на книги.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Получает представление листа с его именем.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|Стиль, примененный к PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Задает стиль, применяемый к PivotTable.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Получает первый pivotTable в коллекции.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Получает сообщение об ошибке запроса с последнего обновления запроса.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Получает запрос 'loaded to' тип объекта.|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Указывает, загружен ли запрос в модель данных.|
||[name](/javascript/api/excel/excel.query#name)|Получает имя запроса.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Получает дату и время последнего обновления запроса.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Получает количество строк, загруженных при последнем обновлении запроса.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Получает количество запросов в книге.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Получает запрос из коллекции на основе его имени.|
||[items](/javascript/api/excel/excel.querycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Возвращает объект, представляющего диапазон, содержащий все иждивенцы ячейки в одной и той же таблице или `WorkbookRangeAreas` в нескольких таблицах.|
||[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Возвращает объект, представляющего диапазон, содержащий все прецеденты ячейки в одной и той же таблице или `WorkbookRangeAreas` в нескольких таблицах.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|Режим обновления связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|Уникальный ID объекта, режим обновления которого был изменен.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Получает тип события.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[обновлено](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Указывает, был ли запрос на обновление успешным.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|Уникальный ID объекта, запрос на обновление которого был завершен.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Получает тип события.|
||[предупреждения](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Массив, содержащий все предупреждения, созданные из запроса на обновление.|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|Получает имя отображения фигуры.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Получает фигуру с ее именем или ИД.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Представляет имя среза, используемое в формуле.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|Стиль, применяемый к срезу.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Задает стиль, примененный к срезу.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Получает стиль по имени.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|Возникает, когда фильтр применяется на определенной таблице.|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|Стиль, примененный к таблице.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|Задает стиль, примененный к таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|Возникает, когда фильтр применяется на любой таблице в книге или в таблице.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|Получает ID таблицы, в которой применяется фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|Получает ID таблицы, которая содержит таблицу.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Удаление нескольких строк из таблицы.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Удаление указанного количества строк из таблицы, начиная с указанного индекса.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|Получает таблицу по имени или ИД.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Возвращает коллекцию связанных типов данных, которые являются частью книги.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Возвращает коллекцию связанных книг.|
||[запросы](/javascript/api/excel/excel.workbook#queries)|Возвращает коллекцию запросов Power Query, которые являются частью книги.|
||[задачи](/javascript/api/excel/excel.workbook#tasks)|Возвращает коллекцию задач, присутствующих в книге.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Указывает, отображается ли область списка полей PivotTable на уровне книги.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Возникает, когда фильтр применяется на определенном таблице.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Возникает при смене состояния защиты таблицы.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Возвращает значение, представляющее этот таблицу, которую можно прочитать в Open Office XML.|
||[задачи](/javascript/api/excel/excel.worksheet#tasks)|Возвращает коллекцию задач, присутствующих в таблице.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Представляет изменение в направлении, в которое будут сдвигаться ячейки в таблице при удалении или вставке ячейки.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Представляет источник триггера события.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Возникает при применении любого фильтра листа в книге.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Возникает при смене состояния защиты таблицы.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Получает ID таблицы, в которой применяется фильтр.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Получает текущее состояние защиты таблицы.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Получает ID таблицы, в которой изменен статус защиты.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
