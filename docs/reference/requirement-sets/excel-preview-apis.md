---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих API JavaScript для Excel.
ms.date: 11/17/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 083741d35d3e881c2e46b186c4e93591bf7f4834
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131768"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Связанные типы данных | Добавляет поддержку для типов данных, подключенных к Excel из внешних источников. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Представления именованных листов | Обеспечивает программное управление представлениями листа на уровне пользователя. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| Задачи | Включение комментариев в задачи, назначенные пользователям. | [Задача](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript для Excel, находящиеся в предварительной версии. Полный список всех API JavaScript для Excel (в том числе API предварительного просмотра и ранее выпущенных API) представлен в статье [все API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[Ассигнтаск (электронная почта: строка)](/javascript/api/excel/excel.comment#assigntask-email-)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве единственного уполномоченного.|
||[Задание ()](/javascript/api/excel/excel.comment#gettask--)|Получает задачу, связанную с данным комментарием.|
||[Жеттаскорнуллобжект ()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Получает задачу, связанную с данным комментарием.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[Ассигнтаск (электронная почта: строка)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве единственного уполномоченного.|
||[Задание ()](/javascript/api/excel/excel.commentreply#gettask--)|Получает задачу, связанную с данным комментарием.|
||[Жеттаскорнуллобжект ()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Получает задачу, связанную с данным комментарием.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[Предоставление dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Имя поставщика данных для связанного типа данных.|
||[ластрефрешед](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Дата и время местного часового пояса с момента открытия книги при последнем обновлении связанного типа данных.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Имя связанного типа данных.|
||[периодикрефрешинтервал](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Частота обновления связанного типа данных (в секундах), если `refreshMode` для параметра задано значение "периодический".|
||[рефрешмоде](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Механизм, с помощью которого извлекаются данные для связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|Уникальный идентификатор связанного типа данных.|
||[суппортедрефрешмодес](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Возвращает массив со всеми режимами обновления, поддерживаемыми связанным типом данных.|
||[Рекуестрефреш ()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Отправляет запрос на обновление связанного типа данных.|
||[Рекуестсетрефрешмоде (Рефрешмоде: Excel. Линкеддататиперефрешмоде)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Отправляет запрос на изменение режима обновления для этого связанного типа данных.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|Уникальный идентификатор нового связанного типа данных.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Получает тип события.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Получает число связанных типов данных в коллекции.|
||[GetItem (ключ: число)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Получает связанный тип данных по идентификатору службы.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Получает связанный тип данных по индексу в коллекции.|
||[getItemOrNullObject (ключ: число)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Получает связанный тип данных по ИДЕНТИФИКАТОРу.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[Рекуестрефрешалл ()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Отправляет запрос на обновление всех связанных типов данных в коллекции.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Активирует это представление листа.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Удаляет представление листа из листа.|
||[дублировать (имя?: строка)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Создает копию этого представления листа.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Получает или задает имя представления листа.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Создает новое представление листа с заданным именем.|
||[Ентертемпорари ()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Создает и активирует новое временное представление листа.|
||[Exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Выполняет выход из текущего активного представления листа.|
||[onactive ()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Получает текущее активное представление листа.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Получает количество просмотров листа на этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Возвращает представление листа с использованием его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Получает представление листа по его индексу в коллекции.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Описание замещающий текст сводной таблицы.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Замещающий текст заголовка сводной таблицы.|
||[Дисплайбланклинеафтереачитем (Display: Boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Указывает, следует ли отображать пустую строку после каждого элемента.|
||[емптицеллтекст](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Текст, автоматически заполняемый в любую пустую ячейку в сводной таблице, если `fillEmptyCells == true` .|
||[филлемптицеллс](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Указывает, следует ли заполнить пустые ячейки в сводной таблице с помощью `emptyCellText` .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий.|
||[пивотстиле](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Стиль, примененный к сводной таблице.|
||[Репеаталлитемлабелс (Репеатлабелс: Boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Задает параметр "повторять все подписи элементов" для всех полей в сводной таблице.|
||[Сетстиле (Style: string \| пивоттаблестиле \| буилтинпивоттаблестиле)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Задает стиль, применяемый к сводной таблице.|
||[шовфиелдхеадерс](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Указывает, отображаются ли в сводной таблице заголовки полей (заголовки полей и раскрывающиеся фильтры).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[рефрешонопен](/javascript/api/excel/excel.pivottable#refreshonopen)|Указывает, обновляется ли Сводная таблица при открытии книги.|
|[Range](/javascript/api/excel/excel.range)|[Влияющие ()](/javascript/api/excel/excel.range#getprecedents--)|Возвращает `WorkbookRangeAreas` объект, представляющий диапазон, содержащий все влияющие ячейки на одном листе или на нескольких листах.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[рефрешмоде](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Режим обновления связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Уникальный идентификатор объекта, режим обновления которого изменился.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Получает тип события.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[обновляется](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Указывает, успешно ли выполнен запрос на обновление.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Уникальный идентификатор объекта, для которого был выполнен запрос на обновление.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Получает тип события.|
||[дефицит](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Массив, содержащий все предупреждения, созданные с помощью запроса на обновление.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Представляет имя среза, используемое в формуле.|
||[слицерстиле](/javascript/api/excel/excel.slicer#slicerstyle)|Стиль, примененный к срезу.|
||[Сетстиле (Style: string \| слицерстиле \| буилтинслицерстиле)](/javascript/api/excel/excel.slicer#setstyle-style-)|Задает стиль, применяемый к срезу.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Возникает, если применен фильтр к указанной таблице.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Стиль, примененный к таблице.|
||[Сетстиле (Style: string \| TableStyle \| буилтинтаблестиле)](/javascript/api/excel/excel.table#setstyle-style-)|Задает стиль, применяемый к таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Возникает, если применен фильтр к любой таблице в книге или листе.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Получает идентификатор таблицы, в которой применяется фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Получает идентификатор листа, содержащего таблицу.|
|[Задача](/javascript/api/excel/excel.task)|[Аддассигни (электронная почта: строка)](/javascript/api/excel/excel.task#addassignee-email-)|Добавляет к задаче уполномоченного.|
||[applyChanges (Таскчанжес: Excel. Таскчанжес)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Применяет указанные изменения к задаче.|
||[уполномоченные](/javascript/api/excel/excel.task#assignees)|Получает пользователей, которым назначена задача.|
||[comment](/javascript/api/excel/excel.task#comment)|Получает комментарий, связанный с задачей.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Получает дату и время выполнения задачи.|
||[хисторирекордс](/javascript/api/excel/excel.task#historyrecords)|Получает записи журнала задачи.|
||[id](/javascript/api/excel/excel.task#id)|Получает идентификатор задачи.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Получает процент завершения задачи.|
||[priority](/javascript/api/excel/excel.task#priority)|Получает приоритет задачи.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Получает дату и время начала задачи.|
||[заголовок](/javascript/api/excel/excel.task#title)|Получает название задачи.|
||[Ремовеаллассигнис ()](/javascript/api/excel/excel.task#removeallassignees--)|Удаляет все уполномоченные из задачи.|
||[Ремовеассигни (электронная почта: строка)](/javascript/api/excel/excel.task#removeassignee-email-)|Удаляет уполномоченного из задачи.|
||[Сетперценткомплете (percentComplete: число)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Изменение завершения задачи.|
||[setPriority (приоритет: число)](/javascript/api/excel/excel.task#setpriority-priority-)|Изменение приоритета задачи.|
||[Сетстартдатеанддуедате (startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Изменение даты начала и срока выполнения задачи.|
||[Сеттитле (Title: строка)](/javascript/api/excel/excel.task#settitle-title-)|Изменяет название задачи.|
|[таскчанжес](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Задает новую дату выполнения для задачи в формате UTC.|
||[емаилстоассигн](/javascript/api/excel/excel.taskchanges#emailstoassign)|Задает адреса электронной почты пользователей, которым назначена задача.|
||[емаилстаунассигн](/javascript/api/excel/excel.taskchanges#emailstounassign)|Задает адреса электронной почты пользователей, для которых требуется отменить назначение задачи.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Задает новый процент завершения для задачи.|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|Задает новый приоритет задачи.|
||[ремовеаллпревиаусассигнис](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Указывает, следует ли удалить из задачи все предыдущие уполномоченные.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Задает новую дату начала для задачи в формате UTC.|
||[заголовок](/javascript/api/excel/excel.taskchanges#title)|Задает новое название задачи.|
|[таскколлектион](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Получает число задач в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Получает задачу с помощью идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Получает задачу по индексу в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Получает задачу с помощью идентификатора.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[таскхисторирекорд](/javascript/api/excel/excel.taskhistoryrecord)|[анчорид](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Представляет идентификатор объекта, к которому привязана задача (например, Комментид для задач, прикрепленных к комментариям).|
||[получателю](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Представляет пользователя, которому назначена задача для типа записи журнала "Assign", или пользователя, чтобы отменить назначение задачи для типа записи журнала "unassign".|
||[аттрибутионусер](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Представляет пользователя, который создал или изменил задачу.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Представляет дату выполнения задачи.|
||[хисторирекордкреатеддате](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Представляет дату создания записи журнала задач.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|Идентификатор записи журнала.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Представляет процент завершения задачи.|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|Представляет приоритет задачи.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Представляет дату начала задачи.|
||[заголовок](/javascript/api/excel/excel.taskhistoryrecord#title)|Представляет название задачи.|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|Представляет тип записи в журнале задач.|
||[ундохисторид](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Представляет свойство TaskHistoryRecord.id, которое было отменено для типа записи истории "Undo".|
|[таскхисторирекордколлектион](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Получает количество записей журнала в коллекции для задачи.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Получает запись журнала задач с помощью индекса в коллекции.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Пользователь](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Представляет отображаемое имя пользователя.|
||[email](/javascript/api/excel/excel.user#email)|Представляет электронный адрес пользователя.|
||[uid](/javascript/api/excel/excel.user#uid)|Представляет уникальный идентификатор пользователя.|
|[Workbook](/javascript/api/excel/excel.workbook)|[линкеддататипес](/javascript/api/excel/excel.workbook#linkeddatatypes)|Возвращает коллекцию связанных типов данных, которые являются частью рабочей книги.|
||[задач](/javascript/api/excel/excel.workbook#tasks)|Возвращает коллекцию задач, присутствующих в книге.|
||[шовпивотфиелдлист](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Указывает, отображается ли область списка полей сводной таблицы на уровне книги.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[намедшитвиевс](/javascript/api/excel/excel.worksheet#namedsheetviews)|Возвращает коллекцию представлений листа, присутствующих на листе.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Возникает, если применен фильтр к указанному листу.|
||[задач](/javascript/api/excel/excel.worksheet#tasks)|Возвращает коллекцию задач, присутствующих на листе.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Возникает при применении любого фильтра листа в книге.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Получает идентификатор листа, в котором применяется фильтр.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
