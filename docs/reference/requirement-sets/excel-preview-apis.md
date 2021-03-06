---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих API JavaScript Excel.
ms.date: 02/24/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0663b6330c402f64e7ed7e8f598a52848bbe1319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505537"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| События с измененной формулой | Отслеживание изменений формул, в том числе источника и типа события, которое вызвало изменение. | [Таблица.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| Связанные типы данных | Добавляет поддержку типов данных, подключенных к Excel из внешних источников. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Именуемые представления листа | Предоставляет программный контроль представлений таблицы для каждого пользователя. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| Задачи | Превратите комментарии в задачи, назначенные пользователям. | [Задача](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript Excel, которые в настоящее время находятся в предварительном просмотре. Полный список всех API JavaScript Excel (включая API предварительного просмотра и ранее выпущенные API) см. во всех API [JavaScript Excel.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask (email: string)](/javascript/api/excel/excel.comment#assigntask-email-)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве единственного назначаемой.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Получает задачу, связанную с этим комментарием.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Получает задачу, связанную с этим комментарием.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (email: string)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве единственного назначаемой.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Получает задачу, связанную с этим комментарием.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Получает задачу, связанную с этим комментарием.|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Адрес ячейки, содержаной измененную формулу.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Представляет предыдущую формулу, прежде чем она была изменена.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Имя поставщика данных для связанного типа данных.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Дата и время локального часового пояса с момента открытия книги при последнем обновлении связанного типа данных.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Имя связанного типа данных.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Частота в секундах, при которой тип связанных данных обновляется, если `refreshMode` установлено "Периодическое".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Механизм получения данных для связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|Уникальный id связанного типа данных.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Возвращает массив со всеми режимами обновления, поддерживаемыми типом связанных данных.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Делает запрос на обновление связанного типа данных.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Делает запрос на изменение режима обновления для этого связанного типа данных.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|Уникальный id нового типа связанных данных.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Получает тип события.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Получает количество связанных типов данных в коллекции.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Получает связанный тип данных по службе id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Получает связанный тип данных по индексу в коллекции.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Получает связанный тип данных по ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Делает запрос на обновление всех связанных типов данных в коллекции.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Активирует это представление листа.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Удаляет представление листа из листа.|
||[дубликат (имя?: строка)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Создает копию этого представления листа.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Получает или задает имя представления листа.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Создает новое представление листа с заданным именем.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Создает и активирует новое временное представление листа.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Выходит из действующего представления листа.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Получает в настоящее время активное представление листа.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Получает количество просмотров листов в этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Получает представление листа с его именем.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Получает представление листа по индексу в коллекции.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|The alt text description of the PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|The alt text title of the PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Задает, следует ли отображать пустую строку после каждого элемента.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Текст, который автоматически заполняется в любую пустую ячейку в PivotTable если `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Указывает, должны ли пустые ячейки в PivotTable заполняться с `emptyCellText` помощью .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Стиль, примененный к PivotTable.|
||[repeatAllItemLabels (repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Задает параметр "Повторите все метки элементов" во всех полях в PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Задает стиль, применяемый к PivotTable.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Указывает, отображаются ли в pivotTable полевые заголовок (подписи полей и отфильтровываемые выпадения).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Указывает, обновляется ли pivotTable при открываемой книге.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Возвращает объект, представляющего диапазон, содержащий все прецеденты ячейки в одной и той же таблице или `WorkbookRangeAreas` в нескольких таблицах.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Режим обновления связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Уникальный id объекта, режим обновления которого был изменен.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Получает тип события.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[обновлено](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Указывает, был ли запрос на обновление успешным.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Уникальный id объекта, запрос на обновление которого был завершен.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Получает тип события.|
||[предупреждения](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Массив, содержащий все предупреждения, созданные из запроса на обновление.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Представляет имя среза, используемое в формуле.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Стиль, примененный к slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Задает стиль, примененный к срезу.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Возникает, если применен фильтр к указанной таблице.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Стиль, примененный к таблице.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Задает стиль, примененный к таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Возникает, если применен фильтр к любой таблице в книге или листе.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Получает id таблицы, в которой применяется фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Получает id таблицы, которая содержит таблицу.|
|[Задача](/javascript/api/excel/excel.task)|[addAssignee (email: string)](/javascript/api/excel/excel.task#addassignee-email-)|Добавляет в задачу назначаемую.|
||[applyChanges(taskChanges: Excel.TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Применяет заданные изменения к задаче.|
||[назначение](/javascript/api/excel/excel.task#assignees)|Получает пользователей, которым назначена задача.|
||[comment](/javascript/api/excel/excel.task#comment)|Получает комментарий, связанный с задачей.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Получает дату и время, когда должна быть поставлена задача.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Получает записи истории задачи.|
||[id](/javascript/api/excel/excel.task#id)|Получает id задачи.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Получает процент выполнения задачи.|
||[priority](/javascript/api/excel/excel.task#priority)|Получает приоритет задачи.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Получает дату и время начала задачи.|
||[заголовок](/javascript/api/excel/excel.task#title)|Получает название задачи.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Удаляет все назначения из задачи.|
||[removeAssignee (email: string)](/javascript/api/excel/excel.task#removeassignee-email-)|Удаляет назначаемую из задачи.|
||[setPercentComplete (percentComplete: number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Изменяет выполнение задачи.|
||[setPriority (приоритет: номер)](/javascript/api/excel/excel.task#setpriority-priority-)|Изменяет приоритет задачи.|
||[setStartDateAndDueDate (startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Изменяет даты начала и срока действия задачи.|
||[setTitle (название: строка)](/javascript/api/excel/excel.task#settitle-title-)|Изменяет название задачи.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Задает новую дату для задачи в часовом поясе UTC.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Задает адреса электронной почты пользователей для назначения задачи.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Задает адреса электронной почты пользователей, чтобы отогнать их от задачи.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Задает новый процент выполнения задачи.|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|Задает новый приоритет для задачи.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Задает, следует ли изменить все предыдущие назначения из задачи.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Задает новую дату начала задачи в часовом поясе UTC.|
||[заголовок](/javascript/api/excel/excel.taskchanges#title)|Задает новое название для задачи.|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Получает количество задач в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Получает задачу с помощью ее id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Получает задачу по индексу в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Получает задачу с помощью ее id.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Представляет ID объекта, на который закреплена задача (например, commentId для задач, присоединенных к комментариям).|
||[assignee](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Представляет пользователя, назначенного для задачи типа записи записи "Назначение" или отозначить из задачи тип записи записи "Unassign".|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Представляет пользователя, создавшего или измениввшего задачу.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Представляет дату окончания задачи.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Представляет дату создания записи истории задач.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|ID для записи истории.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Представляет процент выполнения задачи.|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|Представляет приоритет задачи.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Представляет дату начала задачи.|
||[заголовок](/javascript/api/excel/excel.taskhistoryrecord#title)|Представляет название задачи.|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|Представляет тип записи истории задач.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Представляет свойство TaskHistoryRecord.id, которое было отменено для типа записи записи "Undo".|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Получает количество записей истории в коллекции для задачи.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Получает запись истории задач с помощью ее индекса в коллекции.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[User](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Представляет отображаемое имя пользователя.|
||[email](/javascript/api/excel/excel.user#email)|Представляет электронный адрес пользователя.|
||[uid](/javascript/api/excel/excel.user#uid)|Представляет уникальный ID пользователя.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Возвращает коллекцию связанных типов данных, которые являются частью книги.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Возвращает коллекцию задач, присутствующих в книге.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Указывает, отображается ли область списка полей PivotTable на уровне книги.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Возвращает коллекцию представлений листов, присутствующих в листе.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Возникает, если применен фильтр к указанному листу.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Возникает, когда в этом таблице изменена одна или несколько формул.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Возвращает коллекцию задач, присутствующих в таблице.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Возникает при применении любого фильтра листа в книге.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Возникает, когда одна или несколько формул меняются в любом таблице этой коллекции.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Получает id таблицы, в которой применяется фильтр.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Получает массив объектов FormulaChangedEventDetail, содержащих сведения обо всех измененных формулах.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Получает ID таблицы, в которой изменена формула.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
