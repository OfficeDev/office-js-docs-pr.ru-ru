---
title: Вставка данных в текст в надстройке Outlook
description: Узнайте, как вставить данные в текст сообщения или встречи в надстройке Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: e092a67f8794c2821167ced84bede70a601c77e1
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324956"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a>Вставка данных в текст при создании встречи или сообщения в Outlook

Вы можете использовать асинхронные методы ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) и [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)), чтобы получить тип основного текста и вставить данные в основной текст элемента встречи или сообщения, создаваемых пользователем. Эти асинхронные методы доступны только для надстроек создания. Чтобы использовать эти методы, необходимо настроить манифест для активации надстройки в Outlook, как описано в статье [Создание надстроек Outlook для форм создания](compose-scenario.md).

В Outlook пользователь может создавать сообщения (текстовые, а также в формате HTML и RTF) и встречи (в формате HTML). Перед вставкой всегда необходимо сначала проверить поддерживаемый формат элемента, вызвав метод **getTypeAsync**, так как может понадобиться выполнить дополнительные действия. Значение, которое возвращает метод **getTypeAsync**, зависит от исходного формата элемента, а также от того, поддерживают ли операционная система устройства и узел редактирование в формате HTML (1). Затем соответствующим образом укажите параметр _coercionType_ метода **prependAsync** или **setSelectedDataAsync** (2) для вставки данных, как показано в таблице ниже. Если вы не укажете аргумент, методы **prependAsync** и **setSelectedDataAsync** поведут себя так, как будто данные вставляются в текстовом формате.

<br/>

|**Данные для вставки**|**Формат элемента, возвращенный методом getTypeAsync**|**Необходимый параметр coercionType**|
|:-----|:-----|:-----|
|Текст|Текст (1)|Текст|
|HTML|Текст (1)|Текст (2)|
|Текст|HTML|Текст или HTML|
|HTML|HTML |HTML|

1.  На планшетах и смартфонах метод **getTypeAsync** возвращает **Office.MailboxEnums.BodyType.Text** в формате HTML, если операционная система или узел не поддерживает редактирование элемента, изначально созданного в этом формате.

2.  Если вставляются данные HTML, а метод **getTypeAsync** возвращает текстовый тип, преобразуйте данные в текст и вставьте их, используя **Office.MailboxEnums.BodyType.Text** в качестве _coercionType_. Если просто вставить данные HTML с помощью типа приведения text, узел отобразит HTML-теги в виде текста. Если вы попытаетесь вставить данные HTML, используя **Office.MailboxEnums.BodyType.Html** в качестве _coercionType_, возвратится ошибка.

В дополнение к _coercionType_, как и для большинства асинхронных методов в API JavaScript для Office, **getTypeAsync**, **prependAsync** и **setSelectedDataAsync** принимают другие необязательные входные параметры. Дополнительные сведения об указании дополнительных входных параметров приведены в статье [Передача необязательных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) в [асинхронном программировании в](../develop/asynchronous-programming-in-office-add-ins.md)надстройках Office.


## <a name="insert-data-at-the-current-cursor-position"></a>Вставка данных в текущей позиции курсора


В этом разделе представлен пример кода, который использует **getTypeAsync** для проверки типа текста создаваемого элемента, а затем вызывает метод **setSelectedDataAsync** для вставки данных в текущем положении курсора.

Вы можете передать метод обратного вызова и необязательные входные параметры в **getTypeAsync**. Тогда состояние и результаты будут возвращены в параметре вывода _asyncResult_. Если метод выполнен успешно, вы получите тип текста элемента в свойстве [AsyncResult.value](/javascript/api/office/office.asyncresult#value), значение которого — "text" или "html".

Необходимо передать строку данных как входной параметр метода **setSelectedDataAsync**. В зависимости от типа текста элемента можно указать эту строку в виде текста или HTML соответственно. Как было сказано ранее, при необходимости тип вставляемых данных можно указать в параметре _coercionType_. Кроме того, вы можете предоставить метод обратного вызова и его параметры в качестве дополнительных входных параметров.

Если пользователь не разместил курсор в тексте элемента, **setSelectedDataAsync** вставляет данные в начало текста. Если пользователь выбрал текст в элементе, **setSelectedDataAsync** заменяет выбранный текст указанными вами данными. Обратите внимание, что вызов **setSelectedDataAsync** может завершиться ошибкой, если пользователь одновременно меняет позицию курсора при создании элемента. Максимальное число символов, которые можно вставить за один раз — 1 000 000.

В этом примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.




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


## <a name="insert-data-at-the-beginning-of-the-item-body"></a>Вставка данных в начале текста элемента


Кроме того, с помощью метода **prependAsync** можно вставить данные в начале текста элемента независимо от положения курсора. Помимо точки вставки, методы **prependAsync** и **setSelectedDataAsync** работают одинаково:


- Если вы добавляете HTML-данные в начало текста сообщения, сначала следует проверить тип текста сообщения, чтобы предотвратить вставку HTML-данных в текстовое сообщение.
    
- Предоставьте следующие входные параметры для метода **prependAsync**: строка данных в текстовом формате или формате HTML и, при необходимости, формат вставляемых данных, метод обратного вызова и его параметры.
    
- Максимальное число символов, которые можно вставить в начало за один раз — 1 000 000.
    
Следующий код JavaScript является частью примера надстройки, которая активируется в формах создания встреч и сообщений. Пример вызывает метод **getTypeAsync** для проверки типа текста элемента, вставляет HTML-данные в начало элемента, если это встреча или HTML-сообщение, а в противном случае вставляет данные в текстовом формате.




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


## <a name="see-also"></a>См. также

- [Просмотр и изменение данных элемента в форме создания элементов Outlook](get-and-set-item-data-in-a-compose-form.md)    
- [Просмотр и изменение данных элемента Outlook в формах просмотра и создания](item-data.md)    
- [Создание надстроек Outlook для форм создания](compose-scenario.md)    
- [Асинхронное программирование надстроек Office](../develop/asynchronous-programming-in-office-add-ins.md)    
- [Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook](get-set-or-add-recipients.md)  
- [Просмотр или изменение темы при создании встречи или сообщения в Outlook](get-or-set-the-subject.md)  
- [Просмотр или изменение расположения при создании встречи в Outlook](get-or-set-the-location-of-an-appointment.md) 
- [Просмотр или изменение времени при создании встречи в Outlook](get-or-set-the-time-of-an-appointment.md)
    
