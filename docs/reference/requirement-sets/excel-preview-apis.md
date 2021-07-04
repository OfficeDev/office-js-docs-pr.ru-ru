---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих Excel API JavaScript.
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 39d526f194e1d9e818b8513058d2b414e0bf9673
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290798"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

В следующей таблице приводится краткий сводка [API,](#api-list) а в следующей таблице списка API приводится подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Задачи документа | Превратите комментарии в задачи, назначенные пользователям. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Удостоверения | Управление удостоверениями пользователей, включая имя отображения и адрес электронной почты. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Связанные типы данных | Добавляет поддержку типов данных, подключенных к Excel из внешних источников. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Связанные книги | Управление связями между книгами, включая поддержку обновления и разрыва ссылок на книги. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Стили таблиц | Обеспечивает управление шрифтом, границей, цветом заполнения и другими аспектами стилей таблиц. | [Таблица](/javascript/api/excel/excel.table), [PivotTable](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены Excel API JavaScript, которые в настоящее время находятся в предварительном просмотре. Полный список всех API Excel JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. Excel [API JavaScript.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|Очищает условия фильтрации автофильтра.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteshiftdirection)|Представляет направление (например, вверх или влево), которое остальные ячейки будут смещаться при удалении ячейки или ячейки.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertshiftdirection)|Представляет направление (например, вниз или вправо), в которое будут перенесены существующие ячейки при вставке новой ячейки или ячеек.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assigntask-assignee-)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве ассимилята.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Получает задачу, связанную с этим комментарием.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Получает задачу, связанную с этим комментарием.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|Получает примечание из коллекции на основе его идентификатора.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве единственного назначаемой.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Получает задачу, связанную с потоком ответа на этот комментарий.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Получает задачу, связанную с потоком ответа на этот комментарий.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|Возвращает ответ на примечание, определенное по идентификатору.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|Возвращает условный формат, идентифицированный его ID.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentcomplete)|Указывает процент выполнения задачи.|
||[приоритет](/javascript/api/excel/excel.documenttask#priority)|Указывает приоритет задачи.|
||[назначение](/javascript/api/excel/excel.documenttask#assignees)|Возвращает коллекцию назначений задачи.|
||[изменения](/javascript/api/excel/excel.documenttask#changes)|Получает записи изменений задачи.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Получает комментарий, связанный с задачей.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|Получает последнего пользователя, который выполнил задачу.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|Получает дату и время завершения задачи.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|Получает пользователя, создавшего задачу.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|Получает дату и время создания задачи.|
||[id](/javascript/api/excel/excel.documenttask#id)|Получает ID задачи.|
||[setStartAndDueDateTime (startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|Изменяет даты начала и срока действия задачи.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|Получает или задает дату и время, когда должна начаться и должна быть поставлена задача.|
||[заголовок](/javascript/api/excel/excel.documenttask#title)|Указывает название задачи.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Представляет пользователя, назначенного для задачи для типа записи изменений, или пользователя, не назначенного из задачи `assign` для типа записи `unassign` изменений.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|Представляет пользователя, создавшего или измениввшего задачу.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|Представляет ID того или иного `Comment` изменения `CommentReply` задачи.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|Представляет дату создания и время записи изменения задачи.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|Представляет дату и время задачи в часовом поясе UTC.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID для записи изменения задачи.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|Представляет процент выполнения задачи.|
||[приоритет](/javascript/api/excel/excel.documenttaskchange#priority)|Представляет приоритет задачи.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|Представляет дату и время начала задачи в часовом поясе UTC.|
||[заголовок](/javascript/api/excel/excel.documenttaskchange#title)|Представляет название задачи.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Представляет тип действия записи изменения задачи.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|Представляет `DocumentTaskChange.id` свойство, которое было отменено для типа `undo` записи изменений.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|Получает количество записей изменений в коллекции для задачи.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|Получает запись изменения задачи с помощью индекса в коллекции.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|Получает количество задач в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|Получает задачу с помощью своего ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|Получает задачу по индексу в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|Получает задачу с помощью своего ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|Получает дату и время, когда должна быть поставлена задача.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|Получает дату и время, которые должна начаться задача.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|Получает фигуру с ее именем или ИД.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|Представляет отображаемое имя пользователя.|
||[email](/javascript/api/excel/excel.identity#email)|Представляет электронный адрес пользователя.|
||[id](/javascript/api/excel/excel.identity#id)|Представляет уникальный ID пользователя.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add-assignee-)|Добавляет идентификатор пользователя в коллекцию.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|Удаляет все идентификаторы пользователей из коллекции.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|Возвращает число элементов в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|Получает удостоверение пользователя документа с помощью индекса в коллекции.|
||[items](/javascript/api/excel/excel.identitycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove-assignee-)|Удаляет удостоверение пользователя из коллекции.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayname)|Представляет отображаемое имя пользователя.|
||[email](/javascript/api/excel/excel.identityentity#email)|Представляет электронный адрес пользователя.|
||[id](/javascript/api/excel/excel.identityentity#id)|Представляет уникальный ID пользователя.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Имя поставщика данных для связанного типа данных.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Дата и время локального часового пояса с момента открытия книги при последнем обновлении связанного типа данных.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Имя связанного типа данных.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Частота в секундах, при которой тип связанных данных обновляется, если `refreshMode` установлено "Периодическое".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Механизм получения данных для связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|Уникальный ID связанного типа данных.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Возвращает массив со всеми режимами обновления, поддерживаемыми типом связанных данных.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Делает запрос на обновление связанного типа данных.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Делает запрос на изменение режима обновления для этого связанного типа данных.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|Уникальный ID нового типа связанных данных.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Получает тип события.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Получает количество связанных типов данных в коллекции.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Получает связанный тип данных по ID службы.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Получает связанный тип данных по индексу в коллекции.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Получает связанный тип данных по ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Делает запрос на обновление всех связанных типов данных в коллекции.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breaklinks--)|Делает запрос на разрыв ссылок, указывающих на связанную книгу.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|Исходный URL-адрес, указывающий на связанную книгу.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh--)|Делает запрос на обновление данных, извлеченных из связанной книги.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakalllinks--)|Нарушает все ссылки на связанные книги.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitem-key-)|Получает сведения о связанной книге по URL-адресу.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitemornullobject-key-)|Получает сведения о связанной книге по URL-адресу.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshall--)|Делает запрос на обновление всех ссылок на книги.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbooklinksrefreshmode)|Представляет режим обновления ссылок на книги.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|Получает представление листа с его именем.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Стиль, примененный к PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Задает стиль, применяемый к PivotTable.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|Получает первый pivotTable в коллекции.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getdependents--)|Возвращает объект, представляющего диапазон, содержащий все иждивенцы ячейки в одной и той же таблице или `WorkbookRangeAreas` в нескольких таблицах.|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Возвращает объект, представляющего диапазон, содержащий все прецеденты ячейки в одной и той же таблице или `WorkbookRangeAreas` в нескольких таблицах.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Режим обновления связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Уникальный ID объекта, режим обновления которого был изменен.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Получает тип события.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[обновлено](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Указывает, был ли запрос на обновление успешным.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Уникальный ID объекта, запрос на обновление которого был завершен.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Получает тип события.|
||[предупреждения](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Массив, содержащий все предупреждения, созданные из запроса на обновление.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|Получает фигуру с ее именем или ИД.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Представляет имя среза, используемое в формуле.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Стиль, применяемый к срезу.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Задает стиль, примененный к срезу.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|Получает стиль по имени.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Возникает, когда фильтр применяется на определенной таблице.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Стиль, примененный к таблице.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Задает стиль, примененный к таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Возникает, когда фильтр применяется на любой таблице в книге или в таблице.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Получает ID таблицы, в которой применяется фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Получает ID таблицы, которая содержит таблицу.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|Получает таблицу по имени или ИД.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Возвращает коллекцию связанных типов данных, которые являются частью книги.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedworkbooks)|Возвращает коллекцию связанных книг.|
||[задачи](/javascript/api/excel/excel.workbook#tasks)|Возвращает коллекцию задач, присутствующих в книге.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Указывает, отображается ли область списка полей PivotTable на уровне книги.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Возникает, когда фильтр применяется на определенном таблице.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onprotectionchanged)|Возникает при смене состояния защиты таблицы.|
||[tabId](/javascript/api/excel/excel.worksheet#tabid)|Возвращает значение, представляющее этот таблицу, которую можно прочитать в Open Office XML.|
||[задачи](/javascript/api/excel/excel.worksheet#tasks)|Возвращает коллекцию задач, присутствующих в таблице.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changedirectionstate)|Представляет изменение в направлении, в которое будут сдвигаться ячейки в таблице при удалении или вставке ячейки.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggersource)|Представляет источник триггера события.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Возникает при применении любого фильтра листа в книге.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onprotectionchanged)|Возникает при смене состояния защиты таблицы.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Получает ID таблицы, в которой применяется фильтр.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isprotected)|Получает текущее состояние защиты таблицы.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetid)|Получает ID таблицы, в которой изменен статус защиты.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
