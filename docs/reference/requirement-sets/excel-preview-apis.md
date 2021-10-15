---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих Excel API JavaScript.
ms.date: 10/13/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1c60fa7fe41a9606150b5a83c4d611c97427d1ab
ms.sourcegitcommit: 3b187769e86530334ca83cfdb03c1ecfac2ad9a8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/15/2021
ms.locfileid: "60367476"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

В следующей таблице приводится краткий сводка [API,](#api-list) а в следующей таблице списка API приводится подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Таблицы данных диаграммы | Управление внешним видом, форматированием и видимостью таблиц данных на диаграммах. | [Диаграмма](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Настраиваемые типы данных | Расширение существующих типов Excel, включая поддержку отформатированные номера и веб-изображения. | [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue), [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes), [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes), [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue), [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue), [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue), [StringCellValue](/javascript/api/excel/excel.stringcellvalue), [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue), [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| Ошибки типов пользовательских данных| Объекты ошибки, поддерживают настраиваемые типы данных. | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue), [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue), [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue), [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue), [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue), [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue), [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue), [NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue), [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue), [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue), [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue), [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue), [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue), [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|
| Задачи документа | Превратите комментарии в задачи, назначенные пользователям. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Удостоверения | Управление удостоверениями пользователей, включая имя отображения и адрес электронной почты. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Связанные типы данных | Добавляет поддержку типов данных, подключенных к Excel из внешних источников. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Стили таблиц | Обеспечивает управление шрифтом, границей, цветом заполнения и другими аспектами стилей таблиц. | [Таблица](/javascript/api/excel/excel.table), [PivotTable](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Запросы | Извлечение атрибутов запроса, таких как имя, дата обновления и количество запросов. | [Запрос](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| Защита от таблиц | Запретить неавторизованным пользователям вносить изменения в указанные диапазоны в составе таблицы. | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) [](/javascript/api/excel/excel.alloweditrangecollection) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены Excel API JavaScript, которые в настоящее время находятся в предварительном просмотре. Полный список всех API Excel JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. Excel [API JavaScript.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Объект AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#address)|Указывает диапазон, связанный с объектом.|
||[delete()](/javascript/api/excel/excel.alloweditrange#delete__)|Удаляет этот объект из `AllowEditRangeCollection` .|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#isPasswordProtected)|Указывает, защищен `AllowEditRange` ли пароль.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#pauseProtection_password_)|Приостановка защиты таблиц для данного `AllowEditRange` объекта для пользователя в заданном сеансе.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#setPassword_password_)|Изменяет пароль, связанный с `AllowEditRange` .|
||[заголовок](/javascript/api/excel/excel.alloweditrange#title)|Указывает название объекта.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel. AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#add_title__rangeAddress__options_)|Добавляет объект `AllowEditRange` в коллекцию.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#getCount__)|Возвращает количество объектов `AllowEditRange` в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItem_key_)|Получает объект `AllowEditRange` по его названию.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#getItemAt_index_)|Возвращает объект `AllowEditRange` по индексу в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItemOrNullObject_key_)|Получает объект `AllowEditRange` по его названию.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[pauseProtection (пароль: строка)](/javascript/api/excel/excel.alloweditrangecollection#pauseProtection_password_)|Приостановка защиты от таблиц для всех объектов в коллекции с заданным паролем `AllowEditRange` для пользователя в заданном сеансе.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#password)|Пароль, связанный с `AllowEditRange` .|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#errorSubType)|Представляет тип `BlockedErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.blockederrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.blockederrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[примитивный](/javascript/api/excel/excel.booleancellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.booleancellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.booleancellvalue#type)|Представляет тип этого значения ячейки.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#errorSubType)|Представляет тип `BusyErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.busyerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.busyerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#errorSubType)|Представляет тип `CalcErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.calcerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.calcerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#licenseAddress)|Представляет URL-адрес лицензии или источника, описывая, как можно использовать это свойство.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#licenseText)|Представляет имя лицензии, управляющей этим свойством.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#sourceAddress)|Представляет URL-адрес источника `CellValue` .|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#sourceText)|Представляет имя источника `CellValue` .|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#description)|Представляет свойство описания поставщика, используемое в представлении карты, если не указан логотип.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoSourceAddress)|Представляет URL-адрес, используемый для загрузки изображения, которое будет использоваться в качестве логотипа в представлении карты.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoTargetAddress)|Представляет URL-адрес, который является объектом навигации, если пользователь щелкает элементом логотипа в представлении карты.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве ассимилята.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Получает задачу, связанную с этим комментарием.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Получает задачу, связанную с этим комментарием.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве единственного назначаемой.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Получает задачу, связанную с потоком ответа на этот комментарий.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Получает задачу, связанную с потоком ответа на этот комментарий.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#errorSubType)|Представляет тип `ConnectErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.connecterrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.connecterrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[errorType](/javascript/api/excel/excel.div0errorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.div0errorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.div0errorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.div0errorcellvalue#type)|Представляет тип этого значения ячейки.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[назначение](/javascript/api/excel/excel.documenttask#assignees)|Возвращает коллекцию назначений задачи.|
||[изменения](/javascript/api/excel/excel.documenttask#changes)|Получает записи изменений задачи.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Получает комментарий, связанный с задачей.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|Получает последнего пользователя, который выполнил задачу.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|Получает дату и время завершения задачи.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|Получает пользователя, создавшего задачу.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|Получает дату и время создания задачи.|
||[id](/javascript/api/excel/excel.documenttask#id)|Получает ID задачи.|
||[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Указывает процент выполнения задачи.|
||[приоритет](/javascript/api/excel/excel.documenttask#priority)|Указывает приоритет задачи.|
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
||[приоритет](/javascript/api/excel/excel.documenttaskchange#priority)|Представляет приоритет задачи.|
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
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[примитивный](/javascript/api/excel/excel.doublecellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.doublecellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.doublecellvalue#type)|Представляет тип этого значения ячейки.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[примитивный](/javascript/api/excel/excel.emptycellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.emptycellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.emptycellvalue#type)|Представляет тип этого значения ячейки.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#errorSubType)|Представляет тип `FieldErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.fielderrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.fielderrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#numberFormat)|Возвращает строку формата номеров, которая используется для отображения этого значения.|
||[примитивный](/javascript/api/excel/excel.formattednumbercellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.formattednumbercellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#type)|Представляет тип этого значения ячейки.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.gettingdataerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.gettingdataerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayName)|Представляет отображаемое имя пользователя.|
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
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|Делает запрос на обновление связанного типа данных.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|Делает запрос на изменение режима обновления для этого связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|Уникальный ID связанного типа данных.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|Возвращает массив со всеми режимами обновления, поддерживаемыми типом связанных данных.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|Уникальный ID нового типа связанных данных.|
||[источник](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Получает тип события.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|Получает количество связанных типов данных в коллекции.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|Получает связанный тип данных по ID службы.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|Получает связанный тип данных по индексу в коллекции.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|Получает связанный тип данных по ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|Делает запрос на обновление всех связанных типов данных в коллекции.|
|[NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue)|[errorType](/javascript/api/excel/excel.naerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.naerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.naerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.naerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[errorType](/javascript/api/excel/excel.nameerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.nameerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.nameerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Получает представление листа с его именем.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[errorType](/javascript/api/excel/excel.nullerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.nullerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.nullerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[errorType](/javascript/api/excel/excel.numerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.numerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.numerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.numerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|Стиль, примененный к PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Задает стиль, применяемый к PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#getDataSourceString__)|Возвращает представление строк источника данных для PivotTable.|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#getDataSourceType__)|Получает тип источника данных для PivotTable.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Получает первый pivotTable в коллекции.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Возвращает объект, представляющего диапазон, содержащий все иждивенцы ячейки в одной и той же таблице или `WorkbookRangeAreas` в нескольких таблицах.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[errorSubType](/javascript/api/excel/excel.referrorcellvalue#errorSubType)|Представляет тип `RefErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.referrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.referrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.referrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|Режим обновления связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|Уникальный ID объекта, режим обновления которого был изменен.|
||[источник](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Получает тип события.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[обновлено](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Указывает, был ли запрос на обновление успешным.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|Уникальный ID объекта, запрос на обновление которого был завершен.|
||[источник](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Получает тип события.|
||[предупреждения](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Массив, содержащий все предупреждения, созданные из запроса на обновление.|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|Получает имя отображения фигуры.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Представляет имя среза, используемое в формуле.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Задает стиль, примененный к срезу.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|Стиль, применяемый к срезу.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#errorSubType)|Представляет тип `SpillErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.spillerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.spillerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#spilledColumns)|Представляет количество столбцов, которые будут разливаться, если бы не было #SPILL! ошибка.|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#spilledRows)|Представляет количество строк, которые разлились бы, если бы не было #SPILL! ошибка.|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[примитивный](/javascript/api/excel/excel.stringcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.stringcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.stringcellvalue#type)|Представляет тип этого значения ячейки.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|Возникает, когда фильтр применяется на определенной таблице.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|Задает стиль, примененный к таблице.|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|Стиль, примененный к таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|Возникает, когда фильтр применяется на любой таблице в книге или в таблице.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|Получает ID таблицы, в которой применяется фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|Получает ID таблицы, которая содержит таблицу.|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#errorSubType)|Представляет тип `ValueErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#errorType)|Представляет тип `ErrorCellValue` .|
||[примитивный](/javascript/api/excel/excel.valueerrorcellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.valueerrorcellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#type)|Представляет тип этого значения ячейки.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[примитивный](/javascript/api/excel/excel.valuetypenotavailablecellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#type)|Представляет тип этого значения ячейки.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#address)|Представляет URL-адрес, с которого будет загружено изображение.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#altText)|Представляет альтернативный текст, который можно использовать в сценариях доступности для описания изображения.|
||[атрибуция](/javascript/api/excel/excel.webimagecellvalue#attribution)|Представляет сведения о присвоении для описания исходных и лицензионных требований для использования этого изображения.|
||[примитивный](/javascript/api/excel/excel.webimagecellvalue#primitive)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[primitiveType](/javascript/api/excel/excel.webimagecellvalue#primitiveType)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[поставщик](/javascript/api/excel/excel.webimagecellvalue#provider)|Представляет сведения, описывавшие сущность или физическое лицо, которое предоставило изображение.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#relatedImagesAddress)|Представляет URL-адрес веб-страницы с изображениями, которые считаются связанными с этим `WebImageCellValue` .|
||[type](/javascript/api/excel/excel.webimagecellvalue#type)|Представляет тип этого значения ячейки.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Возвращает коллекцию связанных типов данных, которые являются частью книги.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Указывает, отображается ли область списка полей PivotTable на уровне книги.|
||[задачи](/javascript/api/excel/excel.workbook#tasks)|Возвращает коллекцию задач, присутствующих в книге.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Возникает, когда фильтр применяется на определенном таблице.|
||[задачи](/javascript/api/excel/excel.worksheet#tasks)|Возвращает коллекцию задач, присутствующих в таблице.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Возникает при применении любого фильтра листа в книге.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Получает ID таблицы, в которой применяется фильтр.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#allowEditRanges)|Указывает `AllowEditRangeCollection` найденное в этом документе.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#canPauseProtection)|Указывает, можно ли приостановить защиту для этого таблицы.|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#checkPassword_password_)|Указывает, можно ли использовать пароль для разблокировки защиты таблицы.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#isPasswordProtected)|Указывает, защищен ли лист паролем.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#isPaused)|Указывает, приостановлена ли защита таблицы.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#pauseProtection_password_)|Приостановка защиты таблиц для данного объекта таблицы для пользователя в заданном сеансе.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#resumeProtection__)|Возобновляет защиту таблиц для данного объекта таблицы для пользователя в заданном сеансе.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#setPassword_password_)|Изменяет пароль, связанный с `WorksheetProtection` объектом.|
||[updateOptions (параметры: Excel. WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#updateOptions_options_)|Измените параметры защиты таблиц, связанные с `WorksheetProtection` объектом.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#allowEditRangesChanged)|Указывает, изменились ли `AllowEditRange` какие-либо объекты.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#protectionOptionsChanged)|Указывает, `WorksheetProtectionOptions` изменились ли изменения.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#sheetPasswordChanged)|Указывает, изменился ли пароль таблицы.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
