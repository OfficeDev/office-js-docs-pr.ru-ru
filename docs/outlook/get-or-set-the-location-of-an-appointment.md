---
title: Просмотр или изменение места встречи в надстройке
description: Узнайте, как просмотреть и изменить место проведения встречи в надстройке Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: cc412da5dd64d8e908b86a81b847f6479dbd4a34
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324970"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="fca71-103">Просмотр или изменение расположения при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="fca71-103">Get or set the location when composing an appointment in Outlook</span></span>

<span data-ttu-id="fca71-104">API JavaScript для Office предоставляет свойства и методы для управления расположением встречи, создаваемой пользователем.</span><span class="sxs-lookup"><span data-stu-id="fca71-104">The Office JavaScript API provides properties and methods to manage the location of an appointment that the user is composing.</span></span> <span data-ttu-id="fca71-105">В настоящее время существует два свойства, которые предоставляют место встречи:</span><span class="sxs-lookup"><span data-stu-id="fca71-105">Currently, there are two properties that provide an appointment's location:</span></span>

- <span data-ttu-id="fca71-106">[Item. Location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): базовый API, с помощью которого можно получить и задать расположение.</span><span class="sxs-lookup"><span data-stu-id="fca71-106">[item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Basic API that allows you to get and set the location.</span></span>
- <span data-ttu-id="fca71-107">[Item. енханцедлокатион](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Расширенный API, позволяющий получить и задать расположение, а также указать [тип расположения](/javascript/api/outlook/office.mailboxenums.locationtype).</span><span class="sxs-lookup"><span data-stu-id="fca71-107">[item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Enhanced API that allows you to get and set the location, and includes specifying the [location type](/javascript/api/outlook/office.mailboxenums.locationtype).</span></span> <span data-ttu-id="fca71-108">Тип задается в том `LocationType.Custom` случае, если `item.location`вы задаете расположение с помощью.</span><span class="sxs-lookup"><span data-stu-id="fca71-108">The type is `LocationType.Custom` if you set the location using `item.location`.</span></span>

<span data-ttu-id="fca71-109">В следующей таблице перечислены API расположения и режимы (например, создание или чтение), где они доступны.</span><span class="sxs-lookup"><span data-stu-id="fca71-109">The following table lists the location APIs and the modes (i.e., Compose or Read) where they are available.</span></span>

| <span data-ttu-id="fca71-110">API</span><span class="sxs-lookup"><span data-stu-id="fca71-110">API</span></span> | <span data-ttu-id="fca71-111">Применяемые режимы встреч</span><span class="sxs-lookup"><span data-stu-id="fca71-111">Applicable appointment modes</span></span> |
|---|---|
| [<span data-ttu-id="fca71-112">Item. Location</span><span class="sxs-lookup"><span data-stu-id="fca71-112">item.location</span></span>](/javascript/api/outlook/office.appointmentread#location) | <span data-ttu-id="fca71-113">Участник или чтение</span><span class="sxs-lookup"><span data-stu-id="fca71-113">Attendee/Read</span></span> |
| [<span data-ttu-id="fca71-114">Item. Location. Async</span><span class="sxs-lookup"><span data-stu-id="fca71-114">item.location.getAsync</span></span>](/javascript/api/outlook/office.location#getasync-options--callback-) | <span data-ttu-id="fca71-115">Организатор/создание</span><span class="sxs-lookup"><span data-stu-id="fca71-115">Organizer/Compose</span></span> |
| [<span data-ttu-id="fca71-116">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="fca71-116">item.location.setAsync</span></span>](/javascript/api/outlook/office.location#setasync-location--options--callback-) | <span data-ttu-id="fca71-117">Организатор/создание</span><span class="sxs-lookup"><span data-stu-id="fca71-117">Organizer/Compose</span></span> |
| [<span data-ttu-id="fca71-118">Item. Енханцедлокатион. Async</span><span class="sxs-lookup"><span data-stu-id="fca71-118">item.enhancedLocation.getAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | <span data-ttu-id="fca71-119">Органайзер/создание,</span><span class="sxs-lookup"><span data-stu-id="fca71-119">Organizer/Compose,</span></span><br><span data-ttu-id="fca71-120">Участник или чтение</span><span class="sxs-lookup"><span data-stu-id="fca71-120">Attendee/Read</span></span> |
| [<span data-ttu-id="fca71-121">Item. Енханцедлокатион. addAsync</span><span class="sxs-lookup"><span data-stu-id="fca71-121">item.enhancedLocation.addAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | <span data-ttu-id="fca71-122">Организатор/создание</span><span class="sxs-lookup"><span data-stu-id="fca71-122">Organizer/Compose</span></span> |
| [<span data-ttu-id="fca71-123">Item. Енханцедлокатион. removeAsync</span><span class="sxs-lookup"><span data-stu-id="fca71-123">item.enhancedLocation.removeAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | <span data-ttu-id="fca71-124">Организатор/создание</span><span class="sxs-lookup"><span data-stu-id="fca71-124">Organizer/Compose</span></span> |

<span data-ttu-id="fca71-125">Чтобы использовать методы, доступные только для создания надстроек, настройте манифест надстройки, чтобы активировать надстройку в режиме органайзера или создания.</span><span class="sxs-lookup"><span data-stu-id="fca71-125">To use the methods that are available only to compose add-ins, configure the add-in manifest to activate the add-in in Organizer/Compose mode.</span></span> <span data-ttu-id="fca71-126">Более подробную информацию можно найти в статье [Создание надстроек Outlook для форм создания](compose-scenario.md) .</span><span class="sxs-lookup"><span data-stu-id="fca71-126">See [Create Outlook add-ins for compose forms](compose-scenario.md) for more details.</span></span>

## <a name="use-the-enhancedlocation-api"></a><span data-ttu-id="fca71-127">Использование `enhancedLocation` API</span><span class="sxs-lookup"><span data-stu-id="fca71-127">Use the `enhancedLocation` API</span></span>

<span data-ttu-id="fca71-128">Вы можете использовать `enhancedLocation` API для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="fca71-128">You can use the `enhancedLocation` API to get and set an appointment's location.</span></span> <span data-ttu-id="fca71-129">Поле Location поддерживает несколько расположений, и для каждого местоположения можно задать отображаемое имя, тип и адрес электронной почты комнаты конференц-связи (если это возможно).</span><span class="sxs-lookup"><span data-stu-id="fca71-129">The location field supports multiple locations and, for each location, you can set the display name, type, and conference room email address (if applicable).</span></span> <span data-ttu-id="fca71-130">Поддерживаемые типы расположений представлены в [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) .</span><span class="sxs-lookup"><span data-stu-id="fca71-130">See [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) for supported location types.</span></span>

### <a name="add-location"></a><span data-ttu-id="fca71-131">Добавление расположения</span><span class="sxs-lookup"><span data-stu-id="fca71-131">Add location</span></span>

<span data-ttu-id="fca71-132">В приведенном ниже примере показано, как добавить расположение, вызвав [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) для [Mailbox. Item. енханцедлокатион](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="fca71-132">The following example shows how to add a location by calling [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

```js
var item;
var locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a><span data-ttu-id="fca71-133">Получение расположения</span><span class="sxs-lookup"><span data-stu-id="fca71-133">Get location</span></span>

<span data-ttu-id="fca71-134">В приведенном ниже примере показано, как получить расположение, вызвав метод [Async](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) для [Mailbox. Item. енханцедлокатион](/javascript/api/outlook/office.appointmentread#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="fca71-134">The following example shows how to get the location by calling [getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (place) {
        console.log("Display name: " + place.displayName);
        console.log("Type: " + place.locationIdentifier.type);
        if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
            console.log("Email address: " + place.emailAddress);
        }
    });
}
```

### <a name="remove-location"></a><span data-ttu-id="fca71-135">Удаление расположения</span><span class="sxs-lookup"><span data-stu-id="fca71-135">Remove location</span></span>

<span data-ttu-id="fca71-136">В приведенном ниже примере показано, как удалить расположение, вызвав [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) для [Mailbox. Item. енханцедлокатион](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="fca71-136">The following example shows how to remove the location by calling [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        Office.context.mailbox.item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a><span data-ttu-id="fca71-137">Использование `location` API</span><span class="sxs-lookup"><span data-stu-id="fca71-137">Use the `location` API</span></span>

<span data-ttu-id="fca71-138">Вы можете использовать `location` API для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="fca71-138">You can use the `location` API to get and set an appointment's location.</span></span>

### <a name="get-the-location"></a><span data-ttu-id="fca71-139">Получение места проведения</span><span class="sxs-lookup"><span data-stu-id="fca71-139">Get the location</span></span>

<span data-ttu-id="fca71-140">В этом разделе представлен пример кода, который получает и отображает место проведения создаваемой пользователем встречи.</span><span class="sxs-lookup"><span data-stu-id="fca71-140">This section shows a code sample that gets the location of the appointment that the user is composing, and displays the location.</span></span>

<span data-ttu-id="fca71-141">Чтобы использовать метод `item.location.getAsync`, создайте метод обратного вызова, который проверяет состояние и результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="fca71-141">To use `item.location.getAsync`, provide a callback method that checks for the status and result of the asynchronous call.</span></span> <span data-ttu-id="fca71-142">Вы можете указать все необходимые аргументы метода обратного вызова с помощью необязательного параметра `asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="fca71-142">You can provide any necessary arguments to the callback method through the `asyncContext` optional parameter.</span></span> <span data-ttu-id="fca71-143">Вы можете получать состояние, результаты и любые ошибки, используя выходной параметр `asyncResult` обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fca71-143">You can obtain status, results, and any error using the output parameter `asyncResult` of the callback.</span></span> <span data-ttu-id="fca71-144">Если асинхронный вызов успешно выполнен, вы можете получить место проведения в строковом формате с помощью свойства [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="fca71-144">If the asynchronous call is successful, you can get the location as a string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="set-the-location"></a><span data-ttu-id="fca71-145">Установка места проведения</span><span class="sxs-lookup"><span data-stu-id="fca71-145">Set the location</span></span>

<span data-ttu-id="fca71-146">В этом разделе показан пример кода, который устанавливает место проведения создаваемой пользователем встречи.</span><span class="sxs-lookup"><span data-stu-id="fca71-146">This section shows a code sample that sets the location of the appointment that the user is composing.</span></span>

<span data-ttu-id="fca71-147">Чтобы использовать метод `item.location.setAsync`, укажите строку длиной до 255 символов в параметре data.</span><span class="sxs-lookup"><span data-stu-id="fca71-147">To use `item.location.setAsync`, specify a string of up to 255 characters in the data parameter.</span></span> <span data-ttu-id="fca71-148">При желании вы можете указать метод обратного вызова и его аргументы в параметре `asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="fca71-148">Optionally, you can provide a callback method and any arguments for the callback method in the `asyncContext` parameter.</span></span> <span data-ttu-id="fca71-149">Необходимо проверить состояние, результат и любое сообщение об ошибке в `asyncResult` выходном параметре обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fca71-149">You should check the status, result, and any error message in the `asyncResult` output parameter of the callback.</span></span> <span data-ttu-id="fca71-150">Если асинхронный вызов успешно выполнен, `setAsync` вставляет указанную строку в виде обычного текста, заменяя существующее место проведения.</span><span class="sxs-lookup"><span data-stu-id="fca71-150">If the asynchronous call is successful, `setAsync` inserts the specified location string as plain text, overwriting any existing location for that item.</span></span>

> [!NOTE]
> <span data-ttu-id="fca71-151">Можно задать несколько расположений, используя точку с запятой в качестве разделителя (например, "Конференц-зал A; Конференц-зал B ').</span><span class="sxs-lookup"><span data-stu-id="fca71-151">You can set multiple locations by using a semi-colon as the separator (e.g., 'Conference room A; Conference room B').</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever is appropriate for your scenario,
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## <a name="see-also"></a><span data-ttu-id="fca71-152">См. также</span><span class="sxs-lookup"><span data-stu-id="fca71-152">See also</span></span>

- [<span data-ttu-id="fca71-153">Создание первой надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="fca71-153">Create your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="fca71-154">Асинхронное программирование в случае надстроек Office</span><span class="sxs-lookup"><span data-stu-id="fca71-154">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
