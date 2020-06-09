---
title: Вставка данных в текст в надстройке Outlook
description: Узнайте, как вставить данные в текст сообщения или встречи в надстройке Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: e8100e036d29c13f12aedddd4436cf35569309cf
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609099"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="adfbc-103">Вставка данных в текст при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="adfbc-103">Insert data in the body when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="adfbc-p101">Вы можете использовать асинхронные методы ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) и [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)), чтобы получить тип основного текста и вставить данные в основной текст элемента встречи или сообщения, создаваемых пользователем. Эти асинхронные методы доступны только для надстроек создания. Чтобы использовать эти методы, необходимо настроить манифест для активации надстройки в Outlook, как описано в статье [Создание надстроек Outlook для форм создания](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="adfbc-p101">You can use the asynchronous methods ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) and [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) to get the body type and insert data in the body of an appointment or message item that the user is composing. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately so that Outlook activates your add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="adfbc-p102">В Outlook пользователь может создавать сообщения (текстовые, а также в формате HTML и RTF) и встречи (в формате HTML). Перед вставкой всегда необходимо сначала проверить поддерживаемый формат элемента, вызвав метод **getTypeAsync**, так как может понадобиться выполнить дополнительные действия. Значение, которое возвращает метод **getTypeAsync**, зависит от исходного формата элемента, а также от того, поддерживают ли операционная система устройства и узел редактирование в формате HTML (1). Затем соответствующим образом укажите параметр _coercionType_ метода **prependAsync** или **setSelectedDataAsync** (2) для вставки данных, как показано в таблице ниже. Если вы не укажете аргумент, методы **prependAsync** и **setSelectedDataAsync** поведут себя так, как будто данные вставляются в текстовом формате.</span><span class="sxs-lookup"><span data-stu-id="adfbc-p102">In Outlook, a user can create a message in text, HTML, or Rich Text Format (RTF), and can create an appointment in HTML format. Before inserting, you should always first verify the supported item format by calling **getTypeAsync**, as you may need to take additional steps. The value that **getTypeAsync** returns depends on the original item format, as well as the support of the device operating system and host to editing in HTML format (1). Then set the  _coercionType_ parameter of **prependAsync** or **setSelectedDataAsync** accordingly (2) to insert the data, as shown in the following table. If you don't specify an argument, **prependAsync** and **setSelectedDataAsync** assume the data to insert is in text format.</span></span>

<br/>

|<span data-ttu-id="adfbc-111">**Данные для вставки**</span><span class="sxs-lookup"><span data-stu-id="adfbc-111">**Data to insert**</span></span>|<span data-ttu-id="adfbc-112">**Формат элемента, возвращенный методом getTypeAsync**</span><span class="sxs-lookup"><span data-stu-id="adfbc-112">**Item format returned by getTypeAsync**</span></span>|<span data-ttu-id="adfbc-113">**Необходимый параметр coercionType**</span><span class="sxs-lookup"><span data-stu-id="adfbc-113">**Use this coercionType**</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="adfbc-114">Текст</span><span class="sxs-lookup"><span data-stu-id="adfbc-114">Text</span></span>|<span data-ttu-id="adfbc-115">Текст (1)</span><span class="sxs-lookup"><span data-stu-id="adfbc-115">Text (1)</span></span>|<span data-ttu-id="adfbc-116">Текст</span><span class="sxs-lookup"><span data-stu-id="adfbc-116">Text</span></span>|
|<span data-ttu-id="adfbc-117">HTML</span><span class="sxs-lookup"><span data-stu-id="adfbc-117">HTML</span></span>|<span data-ttu-id="adfbc-118">Текст (1)</span><span class="sxs-lookup"><span data-stu-id="adfbc-118">Text (1)</span></span>|<span data-ttu-id="adfbc-119">Текст (2)</span><span class="sxs-lookup"><span data-stu-id="adfbc-119">Text (2)</span></span>|
|<span data-ttu-id="adfbc-120">Текст</span><span class="sxs-lookup"><span data-stu-id="adfbc-120">Text</span></span>|<span data-ttu-id="adfbc-121">HTML</span><span class="sxs-lookup"><span data-stu-id="adfbc-121">HTML</span></span>|<span data-ttu-id="adfbc-122">Текст или HTML</span><span class="sxs-lookup"><span data-stu-id="adfbc-122">Text/HTML</span></span>|
|<span data-ttu-id="adfbc-123">HTML</span><span class="sxs-lookup"><span data-stu-id="adfbc-123">HTML</span></span>|<span data-ttu-id="adfbc-124">HTML</span><span class="sxs-lookup"><span data-stu-id="adfbc-124">HTML</span></span> |<span data-ttu-id="adfbc-125">HTML</span><span class="sxs-lookup"><span data-stu-id="adfbc-125">HTML</span></span>|

1.  <span data-ttu-id="adfbc-126">На планшетах и смартфонах метод **getTypeAsync** возвращает **Office.MailboxEnums.BodyType.Text** в формате HTML, если операционная система или узел не поддерживает редактирование элемента, изначально созданного в этом формате.</span><span class="sxs-lookup"><span data-stu-id="adfbc-126">On tablets and smartphones, **getTypeAsync** returns **Office.MailboxEnums.BodyType.Text** if the operating system or host does not support editing an item, which was originally created in HTML, in HTML format.</span></span>

2.  <span data-ttu-id="adfbc-p103">Если вставляются данные HTML, а метод **getTypeAsync** возвращает текстовый тип, преобразуйте данные в текст и вставьте их, используя **Office.MailboxEnums.BodyType.Text** в качестве _coercionType_. Если просто вставить данные HTML с помощью типа приведения text, узел отобразит HTML-теги в виде текста. Если вы попытаетесь вставить данные HTML, используя **Office.MailboxEnums.BodyType.Html** в качестве _coercionType_, возвратится ошибка.</span><span class="sxs-lookup"><span data-stu-id="adfbc-p103">If your data to insert is HTML and **getTypeAsync** returns a text type for that item, reorganize your data as text and insert it with **Office.MailboxEnums.BodyType.Text** as _coercionType_. If you simply insert the HTML data with a text coercion type, the host would display the HTML tags as text. If you attempt to insert the HTML data with **Office.MailboxEnums.BodyType.Html** as _coercionType_, you will get an error.</span></span>

<span data-ttu-id="adfbc-p104">В дополнение к _coercionType_, как и для большинства асинхронных методов в API JavaScript для Office, **getTypeAsync**, **prependAsync** и **setSelectedDataAsync** принимают другие необязательные входные параметры. Дополнительные сведения об указании дополнительных входных параметров приведены в статье [Передача необязательных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) в [асинхронном программировании в](../develop/asynchronous-programming-in-office-add-ins.md)надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="adfbc-p104">In addition to  _coercionType_, as with most asynchronous methods in the Office JavaScript API, **getTypeAsync**, **prependAsync** and **setSelectedDataAsync** take other optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="insert-data-at-the-current-cursor-position"></a><span data-ttu-id="adfbc-132">Вставка данных в текущей позиции курсора</span><span class="sxs-lookup"><span data-stu-id="adfbc-132">Insert data at the current cursor position</span></span>


<span data-ttu-id="adfbc-133">В этом разделе представлен пример кода, который использует **getTypeAsync** для проверки типа текста создаваемого элемента, а затем вызывает метод **setSelectedDataAsync** для вставки данных в текущем положении курсора.</span><span class="sxs-lookup"><span data-stu-id="adfbc-133">This section shows a code sample that uses **getTypeAsync** to verify the body type of the item that is being composed, and then uses **setSelectedDataAsync** to insert data in the current cursor location.</span></span>

<span data-ttu-id="adfbc-p105">Вы можете передать метод обратного вызова и необязательные входные параметры в **getTypeAsync**. Тогда состояние и результаты будут возвращены в параметре вывода _asyncResult_. Если метод выполнен успешно, вы получите тип текста элемента в свойстве [AsyncResult.value](/javascript/api/office/office.asyncresult#value), значение которого — "text" или "html".</span><span class="sxs-lookup"><span data-stu-id="adfbc-p105">You can pass a callback method and optional input parameters to **getTypeAsync**, and get any status and results in the  _asyncResult_ output parameter. If the method succeeds, you can get the type of the item body in the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property, which is either "text" or "html".</span></span>

<span data-ttu-id="adfbc-p106">Необходимо передать строку данных как входной параметр метода **setSelectedDataAsync**. В зависимости от типа текста элемента можно указать эту строку в виде текста или HTML соответственно. Как было сказано ранее, при необходимости тип вставляемых данных можно указать в параметре _coercionType_. Кроме того, вы можете предоставить метод обратного вызова и его параметры в качестве дополнительных входных параметров.</span><span class="sxs-lookup"><span data-stu-id="adfbc-p106">You must pass a data string as an input parameter to **setSelectedDataAsync**. Depending on the type of the item body, you can specify this data string in text or HTML format accordingly. As mentioned above, you can optionally specify the type of the data to be inserted in the  _coercionType_ parameter. In addition, you can provide a callback method and any of its parameters as optional input parameters.</span></span>

<span data-ttu-id="adfbc-p107">Если пользователь не разместил курсор в тексте элемента, **setSelectedDataAsync** вставляет данные в начало текста. Если пользователь выбрал текст в элементе, **setSelectedDataAsync** заменяет выбранный текст указанными вами данными. Обратите внимание, что вызов **setSelectedDataAsync** может завершиться ошибкой, если пользователь одновременно меняет позицию курсора при создании элемента. Максимальное число символов, которые можно вставить за один раз — 1 000 000.</span><span class="sxs-lookup"><span data-stu-id="adfbc-p107">If the user hasn't placed the cursor in the item body, **setSelectedDataAsync** inserts the data at the top of the body. If the user has selected text in the item body, **setSelectedDataAsync** replaces the selected text by the data you specify. Note that **setSelectedDataAsync** can fail if the user is simultaneously changing the cursor position while composing the item. The maximum number of characters you can insert at one time is 1,000,000 characters.</span></span>

<span data-ttu-id="adfbc-144">В этом примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="adfbc-144">This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="insert-data-at-the-beginning-of-the-item-body"></a><span data-ttu-id="adfbc-145">Вставка данных в начале текста элемента</span><span class="sxs-lookup"><span data-stu-id="adfbc-145">Insert data at the beginning of the item body</span></span>


<span data-ttu-id="adfbc-p108">Кроме того, с помощью метода **prependAsync** можно вставить данные в начале текста элемента независимо от положения курсора. Помимо точки вставки, методы **prependAsync** и **setSelectedDataAsync** работают одинаково:</span><span class="sxs-lookup"><span data-stu-id="adfbc-p108">Alternatively, you can use **prependAsync** to insert data at the beginning of the item body and disregard the current cursor location. Other than the point of insertion, **prependAsync** and **setSelectedDataAsync** behave in similar ways:</span></span>


- <span data-ttu-id="adfbc-148">Если вы добавляете HTML-данные в начало текста сообщения, сначала следует проверить тип текста сообщения, чтобы предотвратить вставку HTML-данных в текстовое сообщение.</span><span class="sxs-lookup"><span data-stu-id="adfbc-148">If you are prepending HTML data in a message body, you should first check for the type of the message body to avoid prepending HTML data to a message in text format.</span></span>
    
- <span data-ttu-id="adfbc-149">Предоставьте следующие входные параметры для метода **prependAsync**: строка данных в текстовом формате или формате HTML и, при необходимости, формат вставляемых данных, метод обратного вызова и его параметры.</span><span class="sxs-lookup"><span data-stu-id="adfbc-149">Provide the following as input parameters to **prependAsync**: a data string in either text or HTML format, and optionally the format of the data to be inserted, a callback method and any of its parameters.</span></span>
    
- <span data-ttu-id="adfbc-150">Максимальное число символов, которые можно вставить в начало за один раз — 1 000 000.</span><span class="sxs-lookup"><span data-stu-id="adfbc-150">The maximum number of characters you can prepend at one time is 1,000,000 characters.</span></span>
    
<span data-ttu-id="adfbc-p109">Следующий код JavaScript является частью примера надстройки, которая активируется в формах создания встреч и сообщений. Пример вызывает метод **getTypeAsync** для проверки типа текста элемента, вставляет HTML-данные в начало элемента, если это встреча или HTML-сообщение, а в противном случае вставляет данные в текстовом формате.</span><span class="sxs-lookup"><span data-stu-id="adfbc-p109">The following JavaScript code is part of a sample add-in that is activated in compose forms of appointments and messages. The sample calls **getTypeAsync** to verify the type of the item body, inserts HTML data to the top of the item body if the item is an appointment or HTML message, otherwise inserts the data in text format.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="see-also"></a><span data-ttu-id="adfbc-153">См. также</span><span class="sxs-lookup"><span data-stu-id="adfbc-153">See also</span></span>

- [<span data-ttu-id="adfbc-154">Просмотр и изменение данных элемента в форме создания элементов Outlook</span><span class="sxs-lookup"><span data-stu-id="adfbc-154">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="adfbc-155">Просмотр и изменение данных элемента Outlook в формах просмотра и создания</span><span class="sxs-lookup"><span data-stu-id="adfbc-155">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="adfbc-156">Создание надстроек Outlook для форм создания</span><span class="sxs-lookup"><span data-stu-id="adfbc-156">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="adfbc-157">Асинхронное программирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="adfbc-157">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)    
- [<span data-ttu-id="adfbc-158">Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="adfbc-158">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="adfbc-159">Просмотр или изменение темы при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="adfbc-159">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)  
- [<span data-ttu-id="adfbc-160">Просмотр или изменение расположения при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="adfbc-160">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="adfbc-161">Просмотр или изменение времени при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="adfbc-161">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
