---
title: Просмотр и изменение повторения в надстройке Outlook
description: В этой статье показано, как использовать API JavaScript для Office, чтобы просматривать и изменять различные свойства повторения элемента в надстройке Outlook.
ms.date: 08/18/2020
ms.localizationpriority: medium
ms.openlocfilehash: 0b211e72304e22874f847f2231e3a800efaceb4d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154134"
---
# <a name="get-and-set-recurrence"></a>Просмотр и изменение повторения

Иногда требуется создать или обновить повторяющуюся встречу, например еженедельное собрание, посвященное ходу выполнения командного проекта, или ежегодное напоминание о дне рождения. Вы можете использовать Office API JavaScript для управления шаблонами повторения серии встреч в надстройки.

> [!NOTE]
> Поддержка этой функции была представлена в наборе требований 1.7. См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="available-recurrence-patterns"></a>Доступные расписания повторения

Чтобы настроить расписание повторения, необходимо объединить [тип повторения](/javascript/api/outlook/office.mailboxenums.recurrencetype) и его применимые [свойства повторения](/javascript/api/outlook/office.recurrenceproperties) (при наличии).

**Таблица 1. Типы повторений и их применимые свойства**

|Тип повторения|Допустимые свойства повторения|Применение|
|---|---|---|
|`daily`|-&nbsp;[`interval`][interval link]|Встреча проводится через определенный *interval* дней. Пример: встреча проводится каждые **_2_** дня.|
|`weekday`|Отсутствуют.|Встреча повторяется в определенный день недели.|
|`monthly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]|- Встреча проводится в *dayOfMonth* день через определенный *interval* месяцев. Пример: встреча проводится в **_5_** день каждые **_4_** месяца.<br/><br/>- Встреча проводится в *weekNumber* *dayOfWeek* через определенный *interval* месяцев. Пример: встреча проводится в **_третий_** **_четверг_** каждые **_2_** месяца.|
|`weekly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`days`][days link]|Встреча проводится в *days* через определенный *interval* недель. Пример: встреча проводится во **_вторник_ и _четверг_** каждые **_2_** недели.|
|`yearly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]<br/>-&nbsp;[`month`][month link]|- Встреча проводится в *dayOfMonth* день *month* через определенный *interval* лет. Пример: встреча проводится **_7_** **_сентября_** каждые **_4_** года.<br/><br/>- Встреча проводится в *weekNumber* *dayOfWeek* *month* через определенный *interval* лет. Пример: встреча проводится в **_первый_** **_четверг_** **_сентября_** каждые **_2_** года.|

> [!NOTE]
> Вы также можете использовать свойство [`firstDayOfWeek`][firstDayOfWeek link] с типом повторения `weekly`. С указанного дня начинается список дней, отображаемый в диалоговом окне повторения.

## <a name="access-recurrence"></a>Доступ к повторению

Способ доступа к расписанию повторения и действия с ним зависят от того, являетесь ли вы организатором встречи или участником.

**Таблица 2. Применимые состояния встречи**

|Состояние встречи|Возможность изменения повторения|Возможность просмотра повторения|
|---|---|---|
|Организатор встречи — создание ряда|Да ( [`setAsync`][setAsync link] )|Да ( [`getAsync`][getAsync link] )|
|Организатор встречи — создание экземпляра|Нет (`setAsync` возвращает ошибку)|Да ( [`getAsync`][getAsync link] )|
|Участник встречи — чтение ряда|Нет (`setAsync` недоступно)|Да ( [`item.recurrence`][item.recurrence link] )|
|Участник встречи — чтение экземпляра|Нет (`setAsync` недоступно)|Да ( [`item.recurrence`][item.recurrence link] )|
|Приглашение на собрание — чтение ряда|Нет (`setAsync` недоступно)|Да ( [`item.recurrence`][item.recurrence link] )|
|Приглашение на собрание — чтение экземпляра|Нет (`setAsync` недоступно)|Да ( [`item.recurrence`][item.recurrence link] )|

## <a name="set-recurrence-as-the-organizer"></a>Изменение повторения в качестве организатора

Помимо расписания повторения также нужно определить даты и время начала и окончания ряда встреч. Для управления этими сведениями используется объект [`SeriesTime`][SeriesTime link].

Организатор встречи может указать расписание повторения для ряда встреч только в режиме создания. В приведенном ниже примере установлено повторение для ряда встреч с 10:30 до 11:00 (Тихоокеанское время) каждый вторник и четверг со 2 ноября по 2 декабря 2019 г.

```js
var seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

var pattern = {
    "seriesTime": seriesTimeObject,
    "recurrenceType": "weekly",
    "recurrenceProperties": {"interval": 1, "days": ["tue", "thu"]},
    "recurrenceTimeZone": {"name": "Pacific Standard Time"}};

Office.context.mailbox.item.recurrence.setAsync(pattern, callback);

function callback(asyncResult)
{
    console.log(JSON.stringify(asyncResult));
}
```

## <a name="change-recurrence-as-the-organizer"></a>Изменение повторения в качестве организатора

В следующем примере в режиме составить организатор встречи получает объект повторения серии встреч с учетом серии или экземпляра этой серии, а затем задает новую продолжительность повторения.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var recurrencePattern = asyncResult.value;
  recurrencePattern.seriesTime.setDuration(60);
  Office.context.mailbox.item.recurrence.setAsync(recurrencePattern, (asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("failed");
      return;
    }

    console.log("success");
  });
}
```

## <a name="get-recurrence"></a>Просмотр повторения

### <a name="get-recurrence-as-the-organizer"></a>Просмотр повторения в качестве организатора

В приведенном ниже примере организатор встречи просматривает в режиме создания объект повторения ряда встреч с учетом ряда или экземпляра этого ряда.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult){
    var context = asyncResult.context;
    var recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

В приведенном ниже примере показаны результаты вызова `getAsync`, возвращающие повторение для ряда.

> [!NOTE]
> В этом примере `seriesTimeObject` — это заполнитель для JSON, представляющий свойство `recurrence.seriesTime`. Чтобы просмотреть свойства даты и времени повторения, следует использовать методы [`SeriesTime`][SeriesTime link].

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-recurrence-as-an-attendee"></a>Просмотр повторения в качестве участника

В приведенном ниже примере участник встречи может просматривать объект повторения ряда встреч с учетом ряда, экземпляра этого ряда или приглашения на собрание.

```js
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    var recurrence = item.recurrence;
    var seriesId = item.seriesId;

    if (recurrence == null) {
        console.log("Non-recurring item");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

В приведенном ниже примере показано значение свойства `item.recurrence` для ряд встреч.

> [!NOTE]
> В этом примере `seriesTimeObject` — это заполнитель для JSON, представляющий свойство `recurrence.seriesTime`. Чтобы просмотреть свойства даты и времени повторения, следует использовать методы [`SeriesTime`][SeriesTime link].

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-the-recurrence-details"></a>Просмотр сведений повторения

После получения объекта повторение (из обратного вызова `getAsync` или из `item.recurrence`) можно просмотреть определенные свойства повторения. Например, можно просмотреть даты и время начала и окончания ряда с помощью [методов][SeriesTime link] для свойства `recurrence.seriesTime`.

```js
// Get series date and time info
var seriesTime = recurrence.seriesTime;
var startTime = recurrence.seriesTime.getStartTime();
var endTime = recurrence.seriesTime.getEndTime();
var startDate = recurrence.seriesTime.getStartDate();
var endDate = recurrence.seriesTime.getEndDate();
var duration = recurrence.seriesTime.getDuration();

// Get series time zone
var timeZone = recurrence.recurrenceTimeZone;

// Get recurrence properties
var recurrenceProperties = recurrence.recurrenceProperties;

// Get recurrence type
var recurrenceType = recurrence.recurrenceType;
```

## <a name="see-also"></a>См. также

[Событие RecurrenceChanged](/javascript/api/office/office.eventtype)

[getAsync link]: /javascript/api/outlook/office.recurrence#getAsync_options__callback_
[item.recurrence link]: ../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setAsync_recurrencePattern__options__callback_

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayOfMonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayOfWeek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstDayOfWeek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weekNumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
