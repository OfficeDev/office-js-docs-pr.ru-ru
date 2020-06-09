---
title: Просмотр или изменение времени встречи в надстройке Outlook
description: Узнайте, как просмотреть и изменить время начала и окончания встречи в надстройке Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5e02523852584d4b5f1ede9bcd191b9ee16d4c24
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609135"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="5016a-103">Просмотр или изменение времени при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="5016a-103">Get or set the time when composing an appointment in Outlook</span></span>

<span data-ttu-id="5016a-104">API JavaScript для Office предоставляет асинхронные методы ([time. Async](/javascript/api/outlook/office.Time#getasync-options--callback-) и [time. setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) для получения и задания времени начала или окончания встречи, создаваемой пользователем.</span><span class="sxs-lookup"><span data-stu-id="5016a-104">The Office JavaScript API provides asynchronous methods ([Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) and [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) to get and set the start or end time of an appointment that the user is composing.</span></span> <span data-ttu-id="5016a-105">Эти асинхронные методы доступны только для создания надстроек. Чтобы использовать эти методы, убедитесь, что вы правильно настроили манифест надстройки в Outlook для активации надстройки в формах создания, как описано в статье [Создание надстроек Outlook для форм создания](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="5016a-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="5016a-p102">Свойства [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) доступны для встреч в формах создания и чтения. в форме чтения доступ к свойствам можно получить напрямую из родительского объекта, как в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="5016a-p102">The [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:</span></span>

```js
item.start
```

<span data-ttu-id="5016a-108">И в этом примере:</span><span class="sxs-lookup"><span data-stu-id="5016a-108">and in:</span></span>

```js
item.end
```

<span data-ttu-id="5016a-109">Но так как в форме создания и пользователь, и ваша надстройка могут вставлять или изменять сведения о времени одновременно, для получения времени начала и окончания необходимо использовать асинхронный метод **getAsync**, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="5016a-109">But in a compose form, because both the user and your add-in can be inserting or changing the time at the same time, you must use the asynchronous method **getAsync** to get the start or end time, as shown below:</span></span>

```js
item.start.getAsync
```

<span data-ttu-id="5016a-110">И в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="5016a-110">and:</span></span>

```js
item.end.getAsync
```

<span data-ttu-id="5016a-111">Как и в случае с большинством асинхронных методов в API JavaScript для Office, методы SetAsync и- **Async** и **setAsync** принимают необязательные входные параметры.</span><span class="sxs-lookup"><span data-stu-id="5016a-111">As with most asynchronous methods in the Office JavaScript API, **getAsync** and **setAsync** take optional input parameters.</span></span> <span data-ttu-id="5016a-112">Дополнительные сведения об указании последних см. в разделе [Передача дополнительных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) статьи [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="5016a-112">For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-start-or-end-time"></a><span data-ttu-id="5016a-113">Получение времени начала или окончания</span><span class="sxs-lookup"><span data-stu-id="5016a-113">Get the start or end time</span></span>

<span data-ttu-id="5016a-p104">В этом разделе показан пример кода, который получает время начала встречи, создаваемой пользователем, и отображает его. Вы можете использовать тот же код, заменив свойство **start** на **end**, чтобы получить время окончания. В этом примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="5016a-p104">This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.</span></span>


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

<span data-ttu-id="5016a-p105">Чтобы использовать методы **item.start.getAsync** и **item.end.getAsync**, предоставьте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы метода обратного вызова с помощью дополнительного параметра _asyncContext_. Состояние, результаты и сообщения об ошибках можно получить с помощью выходного параметра _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, вы можете получить начальное время как объект **Date** в формате UTC, используя свойство [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="5016a-p105">To use **item.start.getAsync** or **item.end.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the start time as a **Date** object in UTC format using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-start-or-end-time"></a><span data-ttu-id="5016a-121">Установка времени начала или окончания</span><span class="sxs-lookup"><span data-stu-id="5016a-121">Set the start or end time</span></span>

<span data-ttu-id="5016a-p106">В этом разделе показан пример кода, получающий время начало встречи, создаваемой пользователем. Можно использовать тот же код, заменив свойство **start** на **end**, чтобы получить время начала. Обратите внимание, что если у формы создания уже есть время начала, последующая установка времени начала приведет к изменению времени окончания, чтобы сохранить предыдущую длительность встречи. Если у формы создания уже есть время окончания, последующая установка времени окончания приведет к изменению длительности и времени окончания. Если встреча создана как событие на весь день, установки времени начала приведет к смещению времени окончания на 24 часа и отмены выбора параметра события на весь день в форме создания.</span><span class="sxs-lookup"><span data-stu-id="5016a-p106">This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.</span></span>

<span data-ttu-id="5016a-127">Как и в предыдущем примере, здесь предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи.</span><span class="sxs-lookup"><span data-stu-id="5016a-127">Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment.</span></span>

<span data-ttu-id="5016a-p107">Чтобы использовать методы **item.start.setAsync** и **item.end.setAsync**, укажите значение **Date** в формате UTC в параметре _dateTime_. Если вы получаете дату на основе данных, введенных пользователем в клиенте, с помощью [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) можно преобразовать полученное значение в объект **Date** в формате UTC. Можно предоставить необязательный метод обратного вызова и все его аргументы в параметре _asyncContext_. Следует проверить состояние, результат и наличие ошибок в выходном параметре _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, **setAsync** вставляет указанное строку времени начала или окончания как обычный текст, перезаписывая существующее время начала или окончания для этого элемента.</span><span class="sxs-lookup"><span data-stu-id="5016a-p107">To use **item.start.setAsync** or **item.end.setAsync**, specify a **Date** value in UTC in the _dateTime_ parameter. If you get a date based on an input by the user on the client, you can use [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to convert the value to a **Date** object in UTC. You can provide an optional callback method and any arguments for the callback method in the _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified start or end time string as plain text, overwriting any existing start or end time for that item.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="see-also"></a><span data-ttu-id="5016a-133">См. также</span><span class="sxs-lookup"><span data-stu-id="5016a-133">See also</span></span>

- [<span data-ttu-id="5016a-134">Просмотр и изменение данных элемента в форме создания элементов Outlook</span><span class="sxs-lookup"><span data-stu-id="5016a-134">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="5016a-135">Просмотр и изменение данных элемента Outlook в формах просмотра и создания</span><span class="sxs-lookup"><span data-stu-id="5016a-135">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)   
- [<span data-ttu-id="5016a-136">Создание надстроек Outlook для форм создания</span><span class="sxs-lookup"><span data-stu-id="5016a-136">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="5016a-137">Асинхронное программирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5016a-137">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="5016a-138">Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="5016a-138">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="5016a-139">Просмотр или изменение темы при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="5016a-139">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)   
- [<span data-ttu-id="5016a-140">Вставка данных в текст при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="5016a-140">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="5016a-141">Просмотр или изменение расположения при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="5016a-141">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
    
