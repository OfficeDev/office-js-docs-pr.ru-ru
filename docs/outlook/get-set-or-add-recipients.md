---
title: Просмотр и изменение получателей в надстройке Outlook
description: Узнайте, как просмотреть, изменить или добавить получателей сообщения или встречи в надстройке Outlook.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 396f425f639c0d7043154ccfe1ddea16a236f993
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325435"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="feeca-103">Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="feeca-103">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>


<span data-ttu-id="feeca-104">API JavaScript для Office предоставляет асинхронные методы ([Recipients. Async](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients. setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)или [Recipients. addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)), чтобы соответственно получать, задавать и добавлять получателей в форме создания встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="feeca-104">The Office JavaScript API provides asynchronous methods ([Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-), or [Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) to respectively get, set, or add recipients in a compose form of an appointment or message.</span></span> <span data-ttu-id="feeca-105">Эти асинхронные методы доступны только для создания надстроек. Чтобы использовать эти методы, убедитесь, что вы правильно настроили манифест надстройки в Outlook для активации надстройки в формах создания, как описано в статье [Создание надстроек Outlook для форм создания](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="feeca-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="feeca-p102">Некоторые свойства, представляющие получателей в сообщении или встрече, доступны для чтения в формах создания и чтения. Это свойства [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) для встреч, а также свойства [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) для сообщений.</span><span class="sxs-lookup"><span data-stu-id="feeca-p102">Some of the properties that represent recipients in an appointment or message are available for read access in a compose form and in a read form. These properties include  [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for appointments, and [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and  [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for messages.</span></span>

<span data-ttu-id="feeca-108">В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:</span><span class="sxs-lookup"><span data-stu-id="feeca-108">In a read form, you can access the property directly from the parent object, such as:</span></span>

```js
item.cc
```

<span data-ttu-id="feeca-109">Но в форме создания, так как пользователь и надстройка могут одновременно вставлять или изменять получателя, необходимо использовать асинхронный метод `getAsync` для получения этих свойств, как показано в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="feeca-109">But in a compose form, because both the user and your add-in can be inserting or changing a recipient at the same time, you must use the asynchronous method `getAsync` to get these properties, as in the following example:</span></span>


```js
item.cc.getAsync
```

<span data-ttu-id="feeca-110">Эти свойства доступны для записи только в формах создания, но не формах чтения.</span><span class="sxs-lookup"><span data-stu-id="feeca-110">These properties are available for write access in only compose forms and not read forms.</span></span>

<span data-ttu-id="feeca-111">Как и в случае с большинством асинхронных методов в API `getAsync`JavaScript `setAsync`для Office `addAsync` , и использовать необязательные входные параметры.</span><span class="sxs-lookup"><span data-stu-id="feeca-111">As with most asynchronous methods in the JavaScript API for Office, `getAsync`, `setAsync`, and `addAsync` take optional input parameters.</span></span> <span data-ttu-id="feeca-112">Дополнительные сведения об указании последних см. в разделе [Передача дополнительных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) статьи [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="feeca-112">For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-recipients"></a><span data-ttu-id="feeca-113">Извлечение получателей</span><span class="sxs-lookup"><span data-stu-id="feeca-113">Get recipients</span></span>


<span data-ttu-id="feeca-p104">В этом разделе показан пример кода, извлекающий получателей создаваемой встречи или сообщения, и показывающий адреса электронной почты получателей. В примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="feeca-p104">This section shows a code sample that gets the recipients of the appointment or message that is being composed, and displays the email addresses of the recipients. The code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

<span data-ttu-id="feeca-116">В API JavaScript для Office, так как свойства, представляющие получателей встречи ( **optionalAttendees** и **requiredAttendees**), отличаются от тех, которые относятся к сообщению ([BCC](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **CC**и **to**), необходимо сначала использовать свойство [Item. ItemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) , чтобы определить, является ли состав создаваемого элемента встречей или сообщением.</span><span class="sxs-lookup"><span data-stu-id="feeca-116">In the Office JavaScript API, because the properties that represent the recipients of an appointment ( **optionalAttendees** and **requiredAttendees**) are different from those of a message ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **cc**, and **to**), you should first use the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to identify whether the item being composed is an appointment or message.</span></span> <span data-ttu-id="feeca-117">В режиме создания все эти свойства встреч и сообщений являются объектами [Recipients](/javascript/api/outlook/office.Recipients) , поэтому можно применить асинхронный метод, `Recipients.getAsync`чтобы получить соответствующих получателей.</span><span class="sxs-lookup"><span data-stu-id="feeca-117">In compose mode, all these properties of appointments and messages are [Recipients](/javascript/api/outlook/office.Recipients) objects, so you can then apply the asynchronous method, `Recipients.getAsync`, to get the corresponding recipients.</span></span>

<span data-ttu-id="feeca-118">Для использования `getAsync` Предоставьте метод обратного вызова для проверки состояния, результатов и любой ошибки, возвращенной асинхронным `getAsync` вызовом.</span><span class="sxs-lookup"><span data-stu-id="feeca-118">To use `getAsync` provide a callback method to check for the status, results, and any error returned by the asynchronous `getAsync` call.</span></span> <span data-ttu-id="feeca-119">Методу обратного вызова можно передать любые аргументы, используя дополнительный параметр  _asyncContext_.</span><span class="sxs-lookup"><span data-stu-id="feeca-119">You can provide any arguments to the callback method using the optional _asyncContext_ parameter.</span></span> <span data-ttu-id="feeca-120">Метод обратного вызова возвращает выходной параметр  _asyncResult_.</span><span class="sxs-lookup"><span data-stu-id="feeca-120">The callback method returns an _asyncResult_ output parameter.</span></span> <span data-ttu-id="feeca-121">Можно `status` использовать свойства `error` и объекта параметра [asyncResult](/javascript/api/office/office.asyncresult) , чтобы проверить состояние и все сообщения об ошибках асинхронного вызова, а также `value` свойство для получения фактических получателей.</span><span class="sxs-lookup"><span data-stu-id="feeca-121">You can use the `status` and `error` properties of the [AsyncResult](/javascript/api/office/office.asyncresult) parameter object to check for status and any error messages of the asynchronous call, and the `value` property to get the actual recipients.</span></span> <span data-ttu-id="feeca-122">Они представляются как массив объектов [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).</span><span class="sxs-lookup"><span data-stu-id="feeca-122">Recipients are represented as an array of [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) objects.</span></span>

<span data-ttu-id="feeca-123">Обратите внимание на `getAsync` то, что метод является асинхронным, если существуют последующие действия, которые зависят от успешного извлечения получателей, необходимо организовать код для запуска таких действий только в соответствующем методе обратного вызова после успешного завершения асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="feeca-123">Note that because the `getAsync` method is asynchronous, if there are subsequent actions that depend on successfully getting the recipients, you should organize your code to start such actions only in the corresponding callback method when the asynchronous call has successfully completed.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-recipients"></a><span data-ttu-id="feeca-124">Установка получателей</span><span class="sxs-lookup"><span data-stu-id="feeca-124">Set recipients</span></span>


<span data-ttu-id="feeca-125">В этом разделе показан пример кода, устанавливающий получателей встречи или сообщения, создаваемых пользователем.</span><span class="sxs-lookup"><span data-stu-id="feeca-125">This section shows a code sample that sets the recipients of the appointment or message that is being composed by the user.</span></span> <span data-ttu-id="feeca-126">При установке получателей все существующие получателей перезаписываются.</span><span class="sxs-lookup"><span data-stu-id="feeca-126">Setting recipients overwrites any existing recipients.</span></span> <span data-ttu-id="feeca-127">Как и предыдущий пример, извлекающий получателей в форме создания, этот пример предполагает, что надстройка активируется в формах создания встреч и сообщений.</span><span class="sxs-lookup"><span data-stu-id="feeca-127">Similar to the previous example that gets recipients in a compose form, this example assumes that the add-in is activated in compose forms for appointments and messages.</span></span> <span data-ttu-id="feeca-128">В этом примере сначала проверяется, является ли состав созданного элемента встречей или сообщением, чтобы применить асинхронный метод, чтобы применить асинхронный метод `Recipients.setAsync`, в соответствующих свойствах, представляющих получателей встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="feeca-128">This example first verifies if the composed item is an appointment or message, so to apply the asynchronous method, `Recipients.setAsync`, on the appropriate properties that represent recipients of the appointment or message.</span></span>

<span data-ttu-id="feeca-129">При вызове `setAsync`укажите массив в качестве входного аргумента для параметра _Recipients_ в одном из следующих форматов:</span><span class="sxs-lookup"><span data-stu-id="feeca-129">When calling `setAsync`, provide an array as input argument for the  _recipients_ parameter, in one of the following formats:</span></span>


- <span data-ttu-id="feeca-130">массив строк, являющихся SMTP-адресами;</span><span class="sxs-lookup"><span data-stu-id="feeca-130">An array of strings that are SMTP addresses.</span></span>
    
- <span data-ttu-id="feeca-131">массив словарей, каждый из которых содержит отображаемое имя и адрес электронной почты, как показано в следующем примере кода;</span><span class="sxs-lookup"><span data-stu-id="feeca-131">An array of dictionaries, each containing a display name and email address, as shown in the following code sample.</span></span>
    
- <span data-ttu-id="feeca-132">Массив `EmailAddressDetails` объектов, аналогичный элементу, возвращаемому `getAsync` методом.</span><span class="sxs-lookup"><span data-stu-id="feeca-132">An array of `EmailAddressDetails` objects, similar to the one returned by the `getAsync` method.</span></span>
    
<span data-ttu-id="feeca-133">При необходимости можно предоставить метод обратного вызова в качестве аргумента ввода для `setAsync` метода, чтобы убедиться, что код, зависящий от успешной настройки получателей, будет выполняться только в том случае, если это произойдет.</span><span class="sxs-lookup"><span data-stu-id="feeca-133">You can optionally provide a callback method as an input argument to the `setAsync` method, to make sure any code that depends on successfully setting the recipients would execute only when that happens.</span></span> <span data-ttu-id="feeca-134">Кроме того, методу обратного вызова можно передать любые аргументы, используя дополнительный параметр  _asyncContext_.</span><span class="sxs-lookup"><span data-stu-id="feeca-134">You can also provide any arguments for the callback method using the optional _asyncContext_ parameter.</span></span> <span data-ttu-id="feeca-135">При использовании метода обратного вызова можно получить доступ к выходному параметру _asyncResult_ и использовать свойства **Status** и **Error** объекта `AsyncResult` Parameter для проверки состояния и любых сообщений об ошибках асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="feeca-135">If you use a callback method, you can access an _asyncResult_ output parameter, and use the **status** and **error** properties of the `AsyncResult` parameter object to check for status and any error messages of the asynchronous call.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## <a name="add-recipients"></a><span data-ttu-id="feeca-136">Добавление получателей</span><span class="sxs-lookup"><span data-stu-id="feeca-136">Add recipients</span></span>

<span data-ttu-id="feeca-137">Если вы не хотите перезаписать существующие получатели в встрече или сообщении, вместо использования `Recipients.setAsync`можно использовать `Recipients.addAsync` асинхронный метод для добавления получателей.</span><span class="sxs-lookup"><span data-stu-id="feeca-137">If you do not want to overwrite any existing recipients in an appointment or message, instead of using `Recipients.setAsync`, you can use the `Recipients.addAsync` asynchronous method to append recipients.</span></span> <span data-ttu-id="feeca-138">`addAsync`работает так же `setAsync` , как и в случае, если требуется аргумент _Recipients_ .</span><span class="sxs-lookup"><span data-stu-id="feeca-138">`addAsync` works similarly as `setAsync` in that it requires a _recipients_ input argument.</span></span> <span data-ttu-id="feeca-139">При необходимости можно предоставить метод обратного вызова и все аргументы для обратного вызова с помощью параметра asyncContext.</span><span class="sxs-lookup"><span data-stu-id="feeca-139">You can optionally provide a callback method, and any arguments for the callback using the asyncContext parameter.</span></span> <span data-ttu-id="feeca-140">Затем можно проверить состояние, результат и любую ошибку асинхронного `addAsync` вызова, используя выходной параметр _asyncResult_ метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="feeca-140">You can then check the status, result, and any error of the asynchronous `addAsync` call by using the _asyncResult_ output parameter of the callback method.</span></span> <span data-ttu-id="feeca-141">В следующем примере проверяется, является ли состав создаваемого элемента встречей, и добавляет в встречу двух обязательных участников.</span><span class="sxs-lookup"><span data-stu-id="feeca-141">The following example checks if the item being composed is an appointment, and appends two required attendees to the appointment.</span></span>


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## <a name="see-also"></a><span data-ttu-id="feeca-142">См. также</span><span class="sxs-lookup"><span data-stu-id="feeca-142">See also</span></span>

- [<span data-ttu-id="feeca-143">Просмотр и изменение данных элемента в форме создания элементов Outlook</span><span class="sxs-lookup"><span data-stu-id="feeca-143">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)
- [<span data-ttu-id="feeca-144">Просмотр и изменение данных элемента Outlook в формах просмотра и создания</span><span class="sxs-lookup"><span data-stu-id="feeca-144">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)
- [<span data-ttu-id="feeca-145">Создание надстроек Outlook для форм создания</span><span class="sxs-lookup"><span data-stu-id="feeca-145">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="feeca-146">Асинхронное программирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="feeca-146">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="feeca-147">Просмотр или изменение темы при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="feeca-147">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)
- [<span data-ttu-id="feeca-148">Вставка данных в текст при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="feeca-148">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)
- [<span data-ttu-id="feeca-149">Просмотр или изменение расположения при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="feeca-149">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
- [<span data-ttu-id="feeca-150">Просмотр или изменение времени при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="feeca-150">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
