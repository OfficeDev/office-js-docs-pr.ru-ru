---
title: Просмотр и изменение повторения в надстройке Outlook
description: В этой статье показано, как использовать API JavaScript для Office, чтобы просматривать и изменять различные свойства повторения элемента в надстройке Outlook.
ms.date: 08/18/2020
localization_priority: Normal
ms.openlocfilehash: 0b179725677f071fe2ae7baf1c719add5ccd8aa7
ms.sourcegitcommit: e9f23a2857b90a7c17e3152292b548a13a90aa33
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/19/2020
ms.locfileid: "46803746"
---
# <a name="get-and-set-recurrence"></a><span data-ttu-id="128fe-103">Просмотр и изменение повторения</span><span class="sxs-lookup"><span data-stu-id="128fe-103">Get and set recurrence</span></span>

<span data-ttu-id="128fe-104">Иногда требуется создать или обновить повторяющуюся встречу, например еженедельное собрание, посвященное ходу выполнения командного проекта, или ежегодное напоминание о дне рождения.</span><span class="sxs-lookup"><span data-stu-id="128fe-104">Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder.</span></span> <span data-ttu-id="128fe-105">Вы можете использовать API JavaScript для Office для управления шаблонами повторения ряда встреч в надстройке.</span><span class="sxs-lookup"><span data-stu-id="128fe-105">You can use the Office JavaScript API to manage the recurrence patterns of an appointment series in your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="128fe-106">Поддержка этой функции появилась в наборе требований 1,7.</span><span class="sxs-lookup"><span data-stu-id="128fe-106">Support for this feature was introduced in requirement set 1.7.</span></span> <span data-ttu-id="128fe-107">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="128fe-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-recurrence-patterns"></a><span data-ttu-id="128fe-108">Доступные расписания повторения</span><span class="sxs-lookup"><span data-stu-id="128fe-108">Available recurrence patterns</span></span>

<span data-ttu-id="128fe-109">Чтобы настроить расписание повторения, необходимо объединить [тип повторения](/javascript/api/outlook/office.mailboxenums.recurrencetype) и его применимые [свойства повторения](/javascript/api/outlook/office.recurrenceproperties) (при наличии).</span><span class="sxs-lookup"><span data-stu-id="128fe-109">To configure the recurrence pattern, you need to combine the [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (if any).</span></span>

<span data-ttu-id="128fe-110">**Таблица 1. Типы повторений и их применимые свойства**</span><span class="sxs-lookup"><span data-stu-id="128fe-110">**Table 1. Recurrence types and their applicable properties**</span></span>

|<span data-ttu-id="128fe-111">Тип повторения</span><span class="sxs-lookup"><span data-stu-id="128fe-111">Recurrence type</span></span>|<span data-ttu-id="128fe-112">Допустимые свойства повторения</span><span class="sxs-lookup"><span data-stu-id="128fe-112">Valid recurrence properties</span></span>|<span data-ttu-id="128fe-113">Применение</span><span class="sxs-lookup"><span data-stu-id="128fe-113">Usage</span></span>|
|---|---|---|
|`daily`|-&nbsp;[`interval`][interval link]|<span data-ttu-id="128fe-114">Встреча проводится через определенный *interval* дней.</span><span class="sxs-lookup"><span data-stu-id="128fe-114">An appointment occurs every *interval* days.</span></span> <span data-ttu-id="128fe-115">Пример: встреча проводится каждые **_2_** дня.</span><span class="sxs-lookup"><span data-stu-id="128fe-115">Example: An appointment occurs every **_2_** days.</span></span>|
|`weekday`|<span data-ttu-id="128fe-116">Отсутствуют.</span><span class="sxs-lookup"><span data-stu-id="128fe-116">None.</span></span>|<span data-ttu-id="128fe-117">Встреча повторяется в определенный день недели.</span><span class="sxs-lookup"><span data-stu-id="128fe-117">An appointment occurs every weekday.</span></span>|
|`monthly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]|<span data-ttu-id="128fe-118">- Встреча проводится в *dayOfMonth* день через определенный *interval* месяцев.</span><span class="sxs-lookup"><span data-stu-id="128fe-118">- An appointment occurs on day *dayOfMonth* every *interval* months.</span></span> <span data-ttu-id="128fe-119">Пример: встреча проводится в **_5_** день каждые **_4_** месяца.</span><span class="sxs-lookup"><span data-stu-id="128fe-119">Example: An appointment occurs on day **_5_** every **_4_** months.</span></span><br/><br/><span data-ttu-id="128fe-120">- Встреча проводится в *weekNumber* *dayOfWeek* через определенный *interval* месяцев.</span><span class="sxs-lookup"><span data-stu-id="128fe-120">- An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months.</span></span> <span data-ttu-id="128fe-121">Пример: встреча проводится в **_третий_** **_четверг_** каждые **_2_** месяца.</span><span class="sxs-lookup"><span data-stu-id="128fe-121">Example: An appointment occurs on the **_third_** **_Thursday_** every **_2_** months.</span></span>|
|`weekly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`days`][days link]|<span data-ttu-id="128fe-122">Встреча проводится в *days* через определенный *interval* недель.</span><span class="sxs-lookup"><span data-stu-id="128fe-122">An appointment occurs on *days* every *interval* weeks.</span></span> <span data-ttu-id="128fe-123">Пример: встреча проводится во **_вторник_ и _четверг_** каждые **_2_** недели.</span><span class="sxs-lookup"><span data-stu-id="128fe-123">Example: An appointment occurs on **_Tuesday_ and _Thursday_** every **_2_** weeks.</span></span>|
|`yearly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]<br/>-&nbsp;[`month`][month link]|<span data-ttu-id="128fe-124">- Встреча проводится в *dayOfMonth* день *month* через определенный *interval* лет.</span><span class="sxs-lookup"><span data-stu-id="128fe-124">- An appointment occurs on day *dayOfMonth* of *month* every *interval* years.</span></span> <span data-ttu-id="128fe-125">Пример: встреча проводится **_7_** **_сентября_** каждые **_4_** года.</span><span class="sxs-lookup"><span data-stu-id="128fe-125">Example: An appointment occurs on day **_7_** of **_September_** every **_4_** years.</span></span><br/><br/><span data-ttu-id="128fe-126">- Встреча проводится в *weekNumber* *dayOfWeek* *month* через определенный *interval* лет.</span><span class="sxs-lookup"><span data-stu-id="128fe-126">- An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years.</span></span> <span data-ttu-id="128fe-127">Пример: встреча проводится в **_первый_** **_четверг_** **_сентября_** каждые **_2_** года.</span><span class="sxs-lookup"><span data-stu-id="128fe-127">Example: An appointment occurs on the **_first_** **_Thursday_** of **_September_** every **_2_** years.</span></span>|

> [!NOTE]
> <span data-ttu-id="128fe-128">Вы также можете использовать свойство [`firstDayOfWeek`][firstDayOfWeek link] с типом повторения `weekly`.</span><span class="sxs-lookup"><span data-stu-id="128fe-128">You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type.</span></span> <span data-ttu-id="128fe-129">С указанного дня начинается список дней, отображаемый в диалоговом окне повторения.</span><span class="sxs-lookup"><span data-stu-id="128fe-129">The specified day will start the list of days displayed in the recurrence dialog.</span></span>

## <a name="access-recurrence"></a><span data-ttu-id="128fe-130">Доступ к повторению</span><span class="sxs-lookup"><span data-stu-id="128fe-130">Access recurrence</span></span>

<span data-ttu-id="128fe-131">Способ доступа к расписанию повторения и действия с ним зависят от того, являетесь ли вы организатором встречи или участником.</span><span class="sxs-lookup"><span data-stu-id="128fe-131">How you access the recurrence pattern and what you can do with it depends on if you're the appointment organizer or an attendee.</span></span>

<span data-ttu-id="128fe-132">**Таблица 2. Применимые состояния встречи**</span><span class="sxs-lookup"><span data-stu-id="128fe-132">**Table 2. Applicable appointment states**</span></span>

|<span data-ttu-id="128fe-133">Состояние встречи</span><span class="sxs-lookup"><span data-stu-id="128fe-133">Appointment state</span></span>|<span data-ttu-id="128fe-134">Возможность изменения повторения</span><span class="sxs-lookup"><span data-stu-id="128fe-134">Is recurrence editable?</span></span>|<span data-ttu-id="128fe-135">Возможность просмотра повторения</span><span class="sxs-lookup"><span data-stu-id="128fe-135">Is recurrence viewable?</span></span>|
|---|---|---|
|<span data-ttu-id="128fe-136">Организатор встречи — создание ряда</span><span class="sxs-lookup"><span data-stu-id="128fe-136">Appointment organizer - compose series</span></span>|<span data-ttu-id="128fe-137">Да ([`setAsync`][setAsync link])</span><span class="sxs-lookup"><span data-stu-id="128fe-137">Yes ([`setAsync`][setAsync link])</span></span>|<span data-ttu-id="128fe-138">Да ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="128fe-138">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="128fe-139">Организатор встречи — создание экземпляра</span><span class="sxs-lookup"><span data-stu-id="128fe-139">Appointment organizer - compose instance</span></span>|<span data-ttu-id="128fe-140">Нет (`setAsync` возвращает ошибку)</span><span class="sxs-lookup"><span data-stu-id="128fe-140">No (`setAsync` returns an error)</span></span>|<span data-ttu-id="128fe-141">Да ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="128fe-141">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="128fe-142">Участник встречи — чтение ряда</span><span class="sxs-lookup"><span data-stu-id="128fe-142">Appointment attendee - read series</span></span>|<span data-ttu-id="128fe-143">Нет (`setAsync` недоступно)</span><span class="sxs-lookup"><span data-stu-id="128fe-143">No (`setAsync` not available)</span></span>|<span data-ttu-id="128fe-144">Да ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="128fe-144">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="128fe-145">Участник встречи — чтение экземпляра</span><span class="sxs-lookup"><span data-stu-id="128fe-145">Appointment attendee - read instance</span></span>|<span data-ttu-id="128fe-146">Нет (`setAsync` недоступно)</span><span class="sxs-lookup"><span data-stu-id="128fe-146">No (`setAsync` not available)</span></span>|<span data-ttu-id="128fe-147">Да ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="128fe-147">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="128fe-148">Приглашение на собрание — чтение ряда</span><span class="sxs-lookup"><span data-stu-id="128fe-148">Meeting request - read series</span></span>|<span data-ttu-id="128fe-149">Нет (`setAsync` недоступно)</span><span class="sxs-lookup"><span data-stu-id="128fe-149">No (`setAsync` not available)</span></span>|<span data-ttu-id="128fe-150">Да ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="128fe-150">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="128fe-151">Приглашение на собрание — чтение экземпляра</span><span class="sxs-lookup"><span data-stu-id="128fe-151">Meeting request - read instance</span></span>|<span data-ttu-id="128fe-152">Нет (`setAsync` недоступно)</span><span class="sxs-lookup"><span data-stu-id="128fe-152">No (`setAsync` not available)</span></span>|<span data-ttu-id="128fe-153">Да ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="128fe-153">Yes ([`item.recurrence`][item.recurrence link])</span></span>|

## <a name="set-recurrence-as-the-organizer"></a><span data-ttu-id="128fe-154">Изменение повторения в качестве организатора</span><span class="sxs-lookup"><span data-stu-id="128fe-154">Set recurrence as the organizer</span></span>

<span data-ttu-id="128fe-155">Помимо расписания повторения также нужно определить даты и время начала и окончания ряда встреч.</span><span class="sxs-lookup"><span data-stu-id="128fe-155">Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series.</span></span> <span data-ttu-id="128fe-156">Для управления этими сведениями используется объект [`SeriesTime`][SeriesTime link].</span><span class="sxs-lookup"><span data-stu-id="128fe-156">The [`SeriesTime`][SeriesTime link] object is used to manage that information.</span></span>

<span data-ttu-id="128fe-157">Организатор встречи может указать расписание повторения для ряда встреч только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="128fe-157">The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only.</span></span> <span data-ttu-id="128fe-158">В приведенном ниже примере установлено повторение для ряда встреч с 10:30 до 11:00 (Тихоокеанское время) каждый вторник и четверг со 2 ноября по 2 декабря 2019 г.</span><span class="sxs-lookup"><span data-stu-id="128fe-158">In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.</span></span>

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

## <a name="change-recurrence-as-the-organizer"></a><span data-ttu-id="128fe-159">Изменение периодичности в качестве организатора</span><span class="sxs-lookup"><span data-stu-id="128fe-159">Change recurrence as the organizer</span></span>

<span data-ttu-id="128fe-160">В следующем примере организатор встречи получает в режиме создания объект повторения ряда встреч с учетом ряда или экземпляра этого ряда, а затем задает новое значение длительности повторения.</span><span class="sxs-lookup"><span data-stu-id="128fe-160">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series, then sets a new recurrence duration.</span></span>

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

## <a name="get-recurrence"></a><span data-ttu-id="128fe-161">Просмотр повторения</span><span class="sxs-lookup"><span data-stu-id="128fe-161">Get recurrence</span></span>

### <a name="get-recurrence-as-the-organizer"></a><span data-ttu-id="128fe-162">Просмотр повторения в качестве организатора</span><span class="sxs-lookup"><span data-stu-id="128fe-162">Get recurrence as the organizer</span></span>

<span data-ttu-id="128fe-163">В приведенном ниже примере организатор встречи просматривает в режиме создания объект повторения ряда встреч с учетом ряда или экземпляра этого ряда.</span><span class="sxs-lookup"><span data-stu-id="128fe-163">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series.</span></span>

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

<span data-ttu-id="128fe-164">В приведенном ниже примере показаны результаты вызова `getAsync`, возвращающие повторение для ряда.</span><span class="sxs-lookup"><span data-stu-id="128fe-164">The following example shows the results of the `getAsync` call that retrieves the recurrence for a series.</span></span>

> [!NOTE]
> <span data-ttu-id="128fe-165">В этом примере `seriesTimeObject` — это заполнитель для JSON, представляющий свойство `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="128fe-165">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="128fe-166">Чтобы просмотреть свойства даты и времени повторения, следует использовать методы [`SeriesTime`][SeriesTime link].</span><span class="sxs-lookup"><span data-stu-id="128fe-166">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-recurrence-as-an-attendee"></a><span data-ttu-id="128fe-167">Просмотр повторения в качестве участника</span><span class="sxs-lookup"><span data-stu-id="128fe-167">Get recurrence as an attendee</span></span>

<span data-ttu-id="128fe-168">В приведенном ниже примере участник встречи может просматривать объект повторения ряда встреч с учетом ряда, экземпляра этого ряда или приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="128fe-168">In the following example, an appointment attendee can get the recurrence object of an appointment series given the series, an instance of that series, or a meeting request.</span></span>

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

<span data-ttu-id="128fe-169">В приведенном ниже примере показано значение свойства `item.recurrence` для ряд встреч.</span><span class="sxs-lookup"><span data-stu-id="128fe-169">The following example shows the value of the `item.recurrence` property for an appointment series.</span></span>

> [!NOTE]
> <span data-ttu-id="128fe-170">В этом примере `seriesTimeObject` — это заполнитель для JSON, представляющий свойство `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="128fe-170">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="128fe-171">Чтобы просмотреть свойства даты и времени повторения, следует использовать методы [`SeriesTime`][SeriesTime link].</span><span class="sxs-lookup"><span data-stu-id="128fe-171">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-the-recurrence-details"></a><span data-ttu-id="128fe-172">Просмотр сведений повторения</span><span class="sxs-lookup"><span data-stu-id="128fe-172">Get the recurrence details</span></span>

<span data-ttu-id="128fe-173">После получения объекта повторение (из обратного вызова `getAsync` или из `item.recurrence`) можно просмотреть определенные свойства повторения.</span><span class="sxs-lookup"><span data-stu-id="128fe-173">After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence.</span></span> <span data-ttu-id="128fe-174">Например, можно просмотреть даты и время начала и окончания ряда с помощью [методов][SeriesTime link] для свойства `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="128fe-174">For example, you can get the start and end dates and times of the series by using [methods][SeriesTime link] on the `recurrence.seriesTime` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="128fe-175">См. также</span><span class="sxs-lookup"><span data-stu-id="128fe-175">See also</span></span>

[<span data-ttu-id="128fe-176">Событие RecurrenceChanged</span><span class="sxs-lookup"><span data-stu-id="128fe-176">RecurrenceChanged event</span></span>](/javascript/api/office/office.eventtype)

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
