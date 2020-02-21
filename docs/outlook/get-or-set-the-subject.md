---
title: Просмотр или изменение темы в надстройке Outlook
description: Узнайте, как просмотреть и изменить тему сообщения или встречи в надстройке Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: b27f6011b1754fa68a1af87f57034e95fd0d54e0
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166683"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="bf1d5-103">Просмотр или изменение темы при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="bf1d5-103">Get or set the subject when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="bf1d5-p101">API JavaScript для Office предоставляет асинхронные методы ([subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) и [subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)), чтобы получать и задавать тему встречи или сообщения, создаваемого пользователем. Эти методы доступны только для надстроек создания. Чтобы использовать их, необходимо настроить манифест для активации надстройки в формах создания Outlook.</span><span class="sxs-lookup"><span data-stu-id="bf1d5-p101">The JavaScript API for Office provides asynchronous methods ([subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) and [subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) to get and set the subject of an appointment or message that the user is composing. These asynchronous methods are available only to compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms.</span></span>

<span data-ttu-id="bf1d5-p102">Свойство **subject** доступно для чтения в формах создания и формах чтения встреч и сообщений. В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:</span><span class="sxs-lookup"><span data-stu-id="bf1d5-p102">The **subject** property is available for read access in both compose and read forms of appointments and messages. In a read form, you can access the property directly from the parent object, as in:</span></span>

```js
item.subject
```

<span data-ttu-id="bf1d5-108">Но так как в форме создания и пользователь, и ваша надстройка могут вставлять или изменять тему одновременно, для получения темы необходимо использовать асинхронный метод **getAsync**, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="bf1d5-108">But in a compose form, because both the user and your add-in can be inserting or changing the subject at the same time, you must use the asynchronous method **getAsync** to get the subject, as shown below:</span></span>

```js
item.subject.getAsync
```

<span data-ttu-id="bf1d5-109">Свойство **subject** доступно для записи только в формах создания, но не в формах чтения.</span><span class="sxs-lookup"><span data-stu-id="bf1d5-109">The **subject** property is available for write access in only compose forms and not in read forms.</span></span>

<span data-ttu-id="bf1d5-p103">Как и большинство асинхронных методов в API JavaScript для Office, методы **getAsync** и **setAsync** принимают необязательные входные параметры. Дополнительные сведения об указании этих параметров см. в разделе "Передача дополнительных параметров в асинхронные методы" статьи [Асинхронное программирование в надстройках для Office](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bf1d5-p103">As with most asynchronous methods in the JavaScript API for Office, **getAsync** and **setAsync** take optional input parameters. For more information about specifying these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-subject"></a><span data-ttu-id="bf1d5-112">Получение темы</span><span class="sxs-lookup"><span data-stu-id="bf1d5-112">Get the subject</span></span>

<span data-ttu-id="bf1d5-p104">В этом разделе показан пример кода, получающий и отображающий тему создаваемой встречи или сообщения. В примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="bf1d5-p104">This section shows a code sample that gets the subject of the appointment or message that the user is composing, and displays the subject. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

<span data-ttu-id="bf1d5-p105">Чтобы использовать метод **item.subject.getAsync**, предоставьте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы метода обратного вызова с помощью дополнительного параметра  _asyncContext_. Состояние, результаты и сообщения об ошибках можно получить с помощью выходного параметра _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, вы можете получить тему как текстовую строку, используя свойство [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="bf1d5-p105">To use **item.subject.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the subject as a plain text string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-subject"></a><span data-ttu-id="bf1d5-119">Установка темы</span><span class="sxs-lookup"><span data-stu-id="bf1d5-119">Set the subject</span></span>


<span data-ttu-id="bf1d5-p106">В этом разделе показан пример кода, задающий тему создаваемой встречи или сообщения. Как и в предыдущем примере, предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="bf1d5-p106">This section shows a code sample that sets the subject of the appointment or message that the user is composing. Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message.</span></span>

<span data-ttu-id="bf1d5-p107">Чтобы использовать метод **item.subject.setAsync**, укажите строку длиной до 255 символов в параметре data. При необходимости можно предоставить метод обратного вызова и все его аргументы в параметре _asyncContext_. Следует проверить состояние, результат и наличие ошибок в выходном параметре _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, **setAsync** вставляет указанную строку темы как обычный текст, перезаписывая существующую тему этого элемента.</span><span class="sxs-lookup"><span data-stu-id="bf1d5-p107">To use **item.subject.setAsync**, specify a string of up to 255 characters in the data parameter. Optionally, you can provide a callback method and any arguments for the callback method in the  _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified subject string as plain text, overwriting any existing subject for that item.</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
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


## <a name="see-also"></a><span data-ttu-id="bf1d5-126">См. также</span><span class="sxs-lookup"><span data-stu-id="bf1d5-126">See also</span></span>

- [<span data-ttu-id="bf1d5-127">Просмотр и изменение данных элемента в форме создания элементов Outlook</span><span class="sxs-lookup"><span data-stu-id="bf1d5-127">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)   
- [<span data-ttu-id="bf1d5-128">Просмотр и изменение данных элемента Outlook в формах просмотра и создания</span><span class="sxs-lookup"><span data-stu-id="bf1d5-128">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="bf1d5-129">Создание надстроек Outlook для форм создания</span><span class="sxs-lookup"><span data-stu-id="bf1d5-129">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="bf1d5-130">Асинхронное программирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bf1d5-130">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="bf1d5-131">Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="bf1d5-131">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="bf1d5-132">Вставка данных в текст при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="bf1d5-132">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="bf1d5-133">Просмотр или изменение расположения при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="bf1d5-133">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="bf1d5-134">Просмотр или изменение времени при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="bf1d5-134">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
