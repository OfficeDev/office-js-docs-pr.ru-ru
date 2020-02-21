---
title: Просмотр и изменение повторения в надстройке Outlook
description: В этой статье показано, как использовать API JavaScript для Office, чтобы просматривать и изменять различные свойства повторения элемента в надстройке Outlook.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: cc7160ed8fe82a02ced9c03bab181df57e66bb54
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166795"
---
# <a name="get-and-set-recurrence"></a><span data-ttu-id="694b5-103">Просмотр и изменение повторения</span><span class="sxs-lookup"><span data-stu-id="694b5-103">Get and set recurrence</span></span>

<span data-ttu-id="694b5-104">Иногда требуется создать или обновить повторяющуюся встречу, например еженедельное собрание, посвященное ходу выполнения командного проекта, или ежегодное напоминание о дне рождения.</span><span class="sxs-lookup"><span data-stu-id="694b5-104">Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder.</span></span> <span data-ttu-id="694b5-105">Управлять расписаниями повторений для рядов встреч в надстройке можно с помощью API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="694b5-105">You can use the JavaScript API for Office to manage the recurrence patterns of an appointment series in your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="694b5-106">Поддержка этой функции появилась в наборе требований 1,7.</span><span class="sxs-lookup"><span data-stu-id="694b5-106">Support for this feature was introduced in requirement set 1.7.</span></span> <span data-ttu-id="694b5-107">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="694b5-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-recurrence-patterns"></a><span data-ttu-id="694b5-108">Доступные расписания повторения</span><span class="sxs-lookup"><span data-stu-id="694b5-108">Available recurrence patterns</span></span>

<span data-ttu-id="694b5-109">Чтобы настроить расписание повторения, необходимо объединить [тип повторения](/javascript/api/outlook/office.mailboxenums.recurrencetype) и его применимые [свойства повторения](/javascript/api/outlook/office.recurrenceproperties) (при наличии).</span><span class="sxs-lookup"><span data-stu-id="694b5-109">To configure the recurrence pattern, you need to combine the [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (if any).</span></span>

<span data-ttu-id="694b5-110">**Таблица 1. Типы повторений и их применимые свойства**</span><span class="sxs-lookup"><span data-stu-id="694b5-110">**Table 1. Recurrence types and their applicable properties**</span></span>

|<span data-ttu-id="694b5-111">Тип повторения</span><span class="sxs-lookup"><span data-stu-id="694b5-111">Recurrence type</span></span>|<span data-ttu-id="694b5-112">Допустимые свойства повторения</span><span class="sxs-lookup"><span data-stu-id="694b5-112">Valid recurrence properties</span></span>|<span data-ttu-id="694b5-113">Применение</span><span class="sxs-lookup"><span data-stu-id="694b5-113">Usage</span></span>|
|---|---|---|
|`daily`|- [`interval`][interval link]|<span data-ttu-id="694b5-114">Встреча проводится через определенный *interval* дней.</span><span class="sxs-lookup"><span data-stu-id="694b5-114">An appointment occurs every *interval* days.</span></span> <span data-ttu-id="694b5-115">Пример: встреча проводится каждые **_2_** дня.</span><span class="sxs-lookup"><span data-stu-id="694b5-115">Example: An appointment occurs every **_2_** days.</span></span>|
|`weekday`|<span data-ttu-id="694b5-116">Отсутствуют.</span><span class="sxs-lookup"><span data-stu-id="694b5-116">None.</span></span>|<span data-ttu-id="694b5-117">Встреча повторяется в определенный день недели.</span><span class="sxs-lookup"><span data-stu-id="694b5-117">An appointment occurs every weekday.</span></span>|
|`monthly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]|<span data-ttu-id="694b5-118">- Встреча проводится в *dayOfMonth* день через определенный *interval* месяцев.</span><span class="sxs-lookup"><span data-stu-id="694b5-118">- An appointment occurs on day *dayOfMonth* every *interval* months.</span></span> <span data-ttu-id="694b5-119">Пример: встреча проводится в **_5_** день каждые **_4_** месяца.</span><span class="sxs-lookup"><span data-stu-id="694b5-119">Example: An appointment occurs on day **_5_** every **_4_** months.</span></span><br/><br/><span data-ttu-id="694b5-120">- Встреча проводится в *weekNumber* *dayOfWeek* через определенный *interval* месяцев.</span><span class="sxs-lookup"><span data-stu-id="694b5-120">- An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months.</span></span> <span data-ttu-id="694b5-121">Пример: встреча проводится в **_третий_** **_четверг_** каждые **_2_** месяца.</span><span class="sxs-lookup"><span data-stu-id="694b5-121">Example: An appointment occurs on the **_third_** **_Thursday_** every **_2_** months.</span></span>|
|`weekly`|- [`interval`][interval link]<br/>- [`days`][days link]|<span data-ttu-id="694b5-122">Встреча проводится в *days* через определенный *interval* недель.</span><span class="sxs-lookup"><span data-stu-id="694b5-122">An appointment occurs on *days* every *interval* weeks.</span></span> <span data-ttu-id="694b5-123">Пример: встреча проводится во **_вторник_ и _четверг_** каждые **_2_** недели.</span><span class="sxs-lookup"><span data-stu-id="694b5-123">Example: An appointment occurs on **_Tuesday_ and _Thursday_** every **_2_** weeks.</span></span>|
|`yearly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]<br/>- [`month`][month link]|<span data-ttu-id="694b5-124">- Встреча проводится в *dayOfMonth* день *month* через определенный *interval* лет.</span><span class="sxs-lookup"><span data-stu-id="694b5-124">- An appointment occurs on day *dayOfMonth* of *month* every *interval* years.</span></span> <span data-ttu-id="694b5-125">Пример: встреча проводится **_7_** **_сентября_** каждые **_4_** года.</span><span class="sxs-lookup"><span data-stu-id="694b5-125">Example: An appointment occurs on day **_7_** of **_September_** every **_4_** years.</span></span><br/><br/><span data-ttu-id="694b5-126">- Встреча проводится в *weekNumber* *dayOfWeek* *month* через определенный *interval* лет.</span><span class="sxs-lookup"><span data-stu-id="694b5-126">- An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years.</span></span> <span data-ttu-id="694b5-127">Пример: встреча проводится в **_первый_** **_четверг_** **_сентября_** каждые **_2_** года.</span><span class="sxs-lookup"><span data-stu-id="694b5-127">Example: An appointment occurs on the **_first_** **_Thursday_** of **_September_** every **_2_** years.</span></span>|

> [!NOTE]
> <span data-ttu-id="694b5-128">Вы также можете использовать свойство [`firstDayOfWeek`][firstDayOfWeek link] с типом повторения `weekly`.</span><span class="sxs-lookup"><span data-stu-id="694b5-128">You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type.</span></span> <span data-ttu-id="694b5-129">С указанного дня начинается список дней, отображаемый в диалоговом окне повторения.</span><span class="sxs-lookup"><span data-stu-id="694b5-129">The specified day will start the list of days displayed in the recurrence dialog.</span></span>

## <a name="access-recurrence"></a><span data-ttu-id="694b5-130">Доступ к повторению</span><span class="sxs-lookup"><span data-stu-id="694b5-130">Access recurrence</span></span>

<span data-ttu-id="694b5-131">Способ доступа к расписанию повторения и действия с ним зависят от того, являетесь ли вы организатором встречи или участником.</span><span class="sxs-lookup"><span data-stu-id="694b5-131">How you access the recurrence pattern and what you can do with it depends on if you're the appointment organizer or an attendee.</span></span>

<span data-ttu-id="694b5-132">**Таблица 2. Применимые состояния встречи**</span><span class="sxs-lookup"><span data-stu-id="694b5-132">**Table 2. Applicable appointment states**</span></span>

|<span data-ttu-id="694b5-133">Состояние встречи</span><span class="sxs-lookup"><span data-stu-id="694b5-133">Appointment state</span></span>|<span data-ttu-id="694b5-134">Возможность изменения повторения</span><span class="sxs-lookup"><span data-stu-id="694b5-134">Is recurrence editable?</span></span>|<span data-ttu-id="694b5-135">Возможность просмотра повторения</span><span class="sxs-lookup"><span data-stu-id="694b5-135">Is recurrence viewable?</span></span>|
|---|---|---|
|<span data-ttu-id="694b5-136">Организатор встречи — создание ряда</span><span class="sxs-lookup"><span data-stu-id="694b5-136">Appointment organizer - compose series</span></span>|<span data-ttu-id="694b5-137">Да ([`setAsync`][setAsync link])</span><span class="sxs-lookup"><span data-stu-id="694b5-137">Yes ([`setAsync`][setAsync link])</span></span>|<span data-ttu-id="694b5-138">Да ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="694b5-138">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="694b5-139">Организатор встречи — создание экземпляра</span><span class="sxs-lookup"><span data-stu-id="694b5-139">Appointment organizer - compose instance</span></span>|<span data-ttu-id="694b5-140">Нет (`setAsync` возвращает ошибку)</span><span class="sxs-lookup"><span data-stu-id="694b5-140">No (`setAsync` returns an error)</span></span>|<span data-ttu-id="694b5-141">Да ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="694b5-141">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="694b5-142">Участник встречи — чтение ряда</span><span class="sxs-lookup"><span data-stu-id="694b5-142">Appointment attendee - read series</span></span>|<span data-ttu-id="694b5-143">Нет (`setAsync` недоступно)</span><span class="sxs-lookup"><span data-stu-id="694b5-143">No (`setAsync` not available)</span></span>|<span data-ttu-id="694b5-144">Да ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="694b5-144">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="694b5-145">Участник встречи — чтение экземпляра</span><span class="sxs-lookup"><span data-stu-id="694b5-145">Appointment attendee - read instance</span></span>|<span data-ttu-id="694b5-146">Нет (`setAsync` недоступно)</span><span class="sxs-lookup"><span data-stu-id="694b5-146">No (`setAsync` not available)</span></span>|<span data-ttu-id="694b5-147">Да ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="694b5-147">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="694b5-148">Приглашение на собрание — чтение ряда</span><span class="sxs-lookup"><span data-stu-id="694b5-148">Meeting request - read series</span></span>|<span data-ttu-id="694b5-149">Нет (`setAsync` недоступно)</span><span class="sxs-lookup"><span data-stu-id="694b5-149">No (`setAsync` not available)</span></span>|<span data-ttu-id="694b5-150">Да ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="694b5-150">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="694b5-151">Приглашение на собрание — чтение экземпляра</span><span class="sxs-lookup"><span data-stu-id="694b5-151">Meeting request - read instance</span></span>|<span data-ttu-id="694b5-152">Нет (`setAsync` недоступно)</span><span class="sxs-lookup"><span data-stu-id="694b5-152">No (`setAsync` not available)</span></span>|<span data-ttu-id="694b5-153">Да ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="694b5-153">Yes ([`item.recurrence`][item.recurrence link])</span></span>|

## <a name="set-recurrence-as-the-organizer"></a><span data-ttu-id="694b5-154">Изменение повторения в качестве организатора</span><span class="sxs-lookup"><span data-stu-id="694b5-154">Set recurrence as the organizer</span></span>

<span data-ttu-id="694b5-155">Помимо расписания повторения также нужно определить даты и время начала и окончания ряда встреч.</span><span class="sxs-lookup"><span data-stu-id="694b5-155">Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series.</span></span> <span data-ttu-id="694b5-156">Для управления этими сведениями используется объект [`SeriesTime`][SeriesTime link].</span><span class="sxs-lookup"><span data-stu-id="694b5-156">The [`SeriesTime`][SeriesTime link] object is used to manage that information.</span></span>

<span data-ttu-id="694b5-157">Организатор встречи может указать расписание повторения для ряда встреч только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="694b5-157">The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only.</span></span> <span data-ttu-id="694b5-158">В приведенном ниже примере установлено повторение для ряда встреч с 10:30 до 11:00 (Тихоокеанское время) каждый вторник и четверг со 2 ноября по 2 декабря 2019 г.</span><span class="sxs-lookup"><span data-stu-id="694b5-158">In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.</span></span>

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

## <a name="get-recurrence"></a><span data-ttu-id="694b5-159">Просмотр повторения</span><span class="sxs-lookup"><span data-stu-id="694b5-159">Get recurrence</span></span>

### <a name="get-recurrence-as-the-organizer"></a><span data-ttu-id="694b5-160">Просмотр повторения в качестве организатора</span><span class="sxs-lookup"><span data-stu-id="694b5-160">Get recurrence as the organizer</span></span>

<span data-ttu-id="694b5-161">В приведенном ниже примере организатор встречи просматривает в режиме создания объект повторения ряда встреч с учетом ряда или экземпляра этого ряда.</span><span class="sxs-lookup"><span data-stu-id="694b5-161">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series.</span></span>

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

<span data-ttu-id="694b5-162">В приведенном ниже примере показаны результаты вызова `getAsync`, возвращающие повторение для ряда.</span><span class="sxs-lookup"><span data-stu-id="694b5-162">The following example shows the results of the `getAsync` call that retrieves the recurrence for a series.</span></span>

> [!NOTE]
> <span data-ttu-id="694b5-163">В этом примере `seriesTimeObject` — это заполнитель для JSON, представляющий свойство `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="694b5-163">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="694b5-164">Чтобы просмотреть свойства даты и времени повторения, следует использовать методы [`SeriesTime`][SeriesTime link].</span><span class="sxs-lookup"><span data-stu-id="694b5-164">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-recurrence-as-an-attendee"></a><span data-ttu-id="694b5-165">Просмотр повторения в качестве участника</span><span class="sxs-lookup"><span data-stu-id="694b5-165">Get recurrence as an attendee</span></span>

<span data-ttu-id="694b5-166">В приведенном ниже примере участник встречи может просматривать объект повторения ряда встреч с учетом ряда, экземпляра этого ряда или приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="694b5-166">In the following example, an appointment attendee can get the recurrence object of an appointment series given the series, an instance of that series, or a meeting request.</span></span>

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

<span data-ttu-id="694b5-167">В приведенном ниже примере показано значение свойства `item.recurrence` для ряд встреч.</span><span class="sxs-lookup"><span data-stu-id="694b5-167">The following example shows the value of the `item.recurrence` property for an appointment series.</span></span>

> [!NOTE]
> <span data-ttu-id="694b5-168">В этом примере `seriesTimeObject` — это заполнитель для JSON, представляющий свойство `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="694b5-168">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="694b5-169">Чтобы просмотреть свойства даты и времени повторения, следует использовать методы [`SeriesTime`][SeriesTime link].</span><span class="sxs-lookup"><span data-stu-id="694b5-169">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-the-recurrence-details"></a><span data-ttu-id="694b5-170">Просмотр сведений повторения</span><span class="sxs-lookup"><span data-stu-id="694b5-170">Get the recurrence details</span></span>

<span data-ttu-id="694b5-171">После получения объекта повторение (из обратного вызова `getAsync` или из `item.recurrence`) можно просмотреть определенные свойства повторения.</span><span class="sxs-lookup"><span data-stu-id="694b5-171">After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence.</span></span> <span data-ttu-id="694b5-172">Например, можно просмотреть даты и время начала и окончания ряда с помощью [методов][SeriesTime link] для свойства `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="694b5-172">For example, you can get the start and end dates and times of the series by using [methods][SeriesTime link] on the `recurrence.seriesTime` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="694b5-173">См. также</span><span class="sxs-lookup"><span data-stu-id="694b5-173">See also</span></span>

[<span data-ttu-id="694b5-174">Событие RecurrenceChanged</span><span class="sxs-lookup"><span data-stu-id="694b5-174">RecurrenceChanged event</span></span>](/javascript/api/office/office.eventtype)

[getAsync link]: /javascript/api/outlook/office.recurrence#getasync-options--callback-
[item.recurrence link]: ../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setasync-recurrencepattern--options--callback-

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayofmonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayofweek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstdayofweek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weeknumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
