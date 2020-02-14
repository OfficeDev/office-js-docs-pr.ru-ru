---
title: Чтение и запись данных в текущую выделенную область документа или электронной таблицы
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 004f78c4cfa5b5665212cebcdb2d71986398c62a
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950588"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="6be5e-102">Чтение и запись данных в текущую выделенную область документа или электронной таблицы</span><span class="sxs-lookup"><span data-stu-id="6be5e-102">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="6be5e-p101">Объект [Document](/javascript/api/office/office.document) предоставляет методы, с помощью которых можно выполнять операции чтения и записи данных над текущим фрагментом, выделенным пользователем, в документе или электронной таблице. Для этого в объекте **Document** имеются методы **getSelectedDataAsync** и **setSelectedDataAsync**. Кроме того, в данной статье рассказывается, как считывать и записывать данные, а также создавать обработчики событий для обнаружения изменений в выделенном пользователем фрагменте.</span><span class="sxs-lookup"><span data-stu-id="6be5e-p101">The [Document](/javascript/api/office/office.document) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="6be5e-p102">Метод **getSelectedDataAsync** работает только для текущего фрагмента, выделенного пользователем. Если вам необходимо сохранить выделенный фрагмент в документе, чтобы он был доступен для чтения и записи во время последующих сеансов работы надстройки, необходимо добавить привязку с помощью метода [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (или создать привязку с помощью любого метода addFrom объекта [Bindings](/javascript/api/office/office.bindings)). Дополнительные сведения о том, как создать привязку к области в документе, а также о чтении и записи данных через привязку см. в разделе [Привязка к областям в документе или электронной таблице](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="6be5e-p102">The  **getSelectedDataAsync** method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](/javascript/api/office/office.bindings) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="6be5e-109">Чтение выбранных данных</span><span class="sxs-lookup"><span data-stu-id="6be5e-109">Read selected data</span></span>


<span data-ttu-id="6be5e-110">В примере ниже показано, как получить данные из выделенного фрагмента в документе с помощью метода [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="6be5e-110">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="6be5e-p103">В этом примере первый параметр _coercionType_ имеет значение **Office.CoercionType.Text** (вы также можете указать этот параметр, используя строку литерала `"text"`). Это означает, что свойство [value](/javascript/api/office/office.asyncresult#status) объекта [AsyncResult](/javascript/api/office/office.asyncresult), доступного из параметра _asyncResult_ в функции обратного вызова, возвратит **строку**, содержащую выделенный текст в документе. Если вы укажете какой-либо другой тип приведения, то получите другие значения. [Office.CoercionType](/javascript/api/office/office.coerciontype) — это перечисление значений доступных типов приведений. **Office.CoercionType.Text** имеет значение text.</span><span class="sxs-lookup"><span data-stu-id="6be5e-p103">In this example, the first  _coercionType_ parameter is specified as **Office.CoercionType.Text** (you can also specify this parameter by using the literal string `"text"`). This means that the [value](/javascript/api/office/office.asyncresult#status) property of the [AsyncResult](/javascript/api/office/office.asyncresult) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](/javascript/api/office/office.coerciontype) is an enumeration of available coercion type values. **Office.CoercionType.Text** evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="6be5e-p104">**Информация о выборе матричного или табличного coercionType для доступа к данным.** Если предполагается динамический рост данных таблицы с добавлением строк и столбцов, при этом требуется обрабатывать заголовки таблицы, следует использовать табличные данные (указав параметр _coercionType_ метода **getSelectedDataAsync** в виде `"table"` или **Office.CoercionType.Table**). Добавление строк и столбцов в структуре данных поддерживается как табличными, так и матричными данными, но присоединение строк и столбцов поддерживается только табличными данными. В отсутствие необходимости добавления строк и столбцов, если при этом не требуется использовать заголовки, следует выбрать матричные данные (указав параметр _coercionType_ метода **getSelecteDataAsync** в виде `"matrix"` или **Office.CoercionType.Matrix**), что позволяет использовать упрощенный способ взаимодействия с данными.</span><span class="sxs-lookup"><span data-stu-id="6be5e-p104">**When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the **getSelectedDataAsync** method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of **getSelecteDataAsync** method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="6be5e-p105">Анонимная функция, которая передается в функцию в качестве второго параметра _callback_, выполняется после завершения операции **getSelectedDataAsync**. При вызове функции передается один параметр _asyncResult_, который содержит результат и сведения о состоянии вызова. Если вызов завершается с ошибкой, свойство [error](/javascript/api/office/office.asyncresult#asynccontext) объекта **AsyncResult** предоставляет доступ к объекту [Error](/javascript/api/office/office.error). Вы можете проверить значение свойств [Error.name](/javascript/api/office/office.error#name) и [Error.message](/javascript/api/office/office.error#message), чтобы определить, почему операция завершилась с ошибкой. В противном случае будет отображен выделенный в документе текст.</span><span class="sxs-lookup"><span data-stu-id="6be5e-p105">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the **getSelectedDataAsync** operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](/javascript/api/office/office.asyncresult#asynccontext) property of the **AsyncResult** object provides access to the [Error](/javascript/api/office/office.error) object. You can check the value of the [Error.name](/javascript/api/office/office.error#name) and [Error.message](/javascript/api/office/office.error#message) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="6be5e-p106">Свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#error) используется в выражении **if** для проверки того, успешно ли выполнен вызов. [Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) — это перечисление доступных значений свойства **AsyncResult.status**. **Office.AsyncResultStatus.Failed** имеет значение failed (и его можно указать в виде строки литералов).</span><span class="sxs-lookup"><span data-stu-id="6be5e-p106">The [AsyncResult.status](/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) is an enumeration of available **AsyncResult.status** property values. **Office.AsyncResultStatus.Failed** evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="6be5e-128">Запись данных в выделение</span><span class="sxs-lookup"><span data-stu-id="6be5e-128">Write data to the selection</span></span>


<span data-ttu-id="6be5e-129">В следующем примере показано, как записать в выделение строку "Hello World!".</span><span class="sxs-lookup"><span data-stu-id="6be5e-129">The following example shows how to set the selection to show "Hello World!".</span></span>


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

<span data-ttu-id="6be5e-p107">Передача в параметре _data_ других типов объектов может привести к разным результатам. Результат зависит от текущего выделения в документе, от ведущего приложения, а также от возможности приведения переданных данных применительно к текущему выделению.</span><span class="sxs-lookup"><span data-stu-id="6be5e-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="6be5e-p108">Анонимная функция, которая передается в метод [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) в качестве параметра _callback_, выполняется после завершения асинхронного вызова. При записи данных в выделенный фрагмент с помощью метода **setSelectedDataAsync** параметр _asyncResult_ обратного вызова предоставляет доступ только к сведениям о состоянии вызова и к объекту [Error](/javascript/api/office/office.error) в случае сбоя вызова.</span><span class="sxs-lookup"><span data-stu-id="6be5e-p108">The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the **setSelectedDataAsync** method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](/javascript/api/office/office.error) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="6be5e-134">Начиная с выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel в Интернете, вы можете [задать форматирование при записи таблицы в текущую выделенную область](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="6be5e-134">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel on the web, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="6be5e-135">Обнаружение изменений в выделенной области</span><span class="sxs-lookup"><span data-stu-id="6be5e-135">Detect changes in the selection</span></span>


<span data-ttu-id="6be5e-136">В примере ниже показано, как определять изменения в выделенном фрагменте, используя метод [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) для добавления обработчика события [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) в документе.</span><span class="sxs-lookup"><span data-stu-id="6be5e-136">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event on the document.</span></span>


```js
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){}
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

<span data-ttu-id="6be5e-p109">Первый параметр _eventType_ задает имя события для подписки. Передача строки `"documentSelectionChanged"` для этого параметра эквивалентна передаче типа события **Office.EventType.DocumentSelectionChanged** перечисления [Office.EventType](/javascript/api/office/office.eventtype).</span><span class="sxs-lookup"><span data-stu-id="6be5e-p109">The first  _eventType_ parameter specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the **Office.EventType.DocumentSelectionChanged** event type of the [Office.EventType](/javascript/api/office/office.eventtype) enumeration.</span></span>

<span data-ttu-id="6be5e-p110">Анонимная функция `myHander()`, передаваемая в эту функцию в качестве второго параметра _handler_, представляет собой обработчик событий, который выполняется при изменении выделенного фрагмента в документе. При вызове этой функции передается единственный параметр _eventArgs_, который после завершения асинхронной операции будет содержать ссылку на объект [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs). Вы можете использовать свойство [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) для доступа к документу, создавшему событие.</span><span class="sxs-lookup"><span data-stu-id="6be5e-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="6be5e-p111">Вы можете добавить несколько обработчиков событий для данного события, снова вызвав метод **addHandlerAsync** и передав дополнительную функцию обработчика события для параметра _handler_. Это будет работать правильно, поскольку имя каждой функции обработчика событий уникально.</span><span class="sxs-lookup"><span data-stu-id="6be5e-p111">You can add multiple event handlers for a given event by calling the  **addHandlerAsync** method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="6be5e-144">Отключение обнаружения изменений в выделенной области</span><span class="sxs-lookup"><span data-stu-id="6be5e-144">Stop detecting changes in the selection</span></span>


<span data-ttu-id="6be5e-145">В примере ниже показано, как остановить прослушивание события [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs), вызвав метод [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="6be5e-145">The following example shows how to stop listening to the [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event by calling the [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="6be5e-146">Имя функции `myHandler`, передаваемое в качестве второго параметра _handler_, задает обработчик событий, который будет удален из события **SelectionChanged**.</span><span class="sxs-lookup"><span data-stu-id="6be5e-146">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the **SelectionChanged** event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="6be5e-147">Если необязательный параметр _handler_ при вызове метода **removeHandlerAsync** не указывается, то все обработчики событий для указанного объекта _eventType_ будут удалены.</span><span class="sxs-lookup"><span data-stu-id="6be5e-147">If the optional  _handler_ parameter is omitted when the **removeHandlerAsync** method is called, all event handlers for the specified _eventType_ will be removed.</span></span>
