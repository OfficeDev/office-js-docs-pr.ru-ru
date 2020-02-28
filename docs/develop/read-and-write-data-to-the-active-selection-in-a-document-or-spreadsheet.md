---
title: Чтение и запись данных в текущую выделенную область документа или электронной таблицы
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 039631e935d2ff6fadb4eab9d99df73ac30dae4d
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325005"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="0e016-102">Чтение и запись данных в текущую выделенную область документа или электронной таблицы</span><span class="sxs-lookup"><span data-stu-id="0e016-102">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="0e016-p101">Объект [Document](/javascript/api/office/office.document) предоставляет методы, позволяющие считывать и записывать выделенный пользователем фрагмент в документе или электронной таблице. Для этого `Document` объект предоставляет методы `getSelectedDataAsync` и. `setSelectedDataAsync` В этом разделе также описывается, как читать, записывать и создавать обработчики событий для обнаружения изменений в выделении пользователя.</span><span class="sxs-lookup"><span data-stu-id="0e016-p101">The [Document](/javascript/api/office/office.document) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet. To do that, the `Document` object provides the `getSelectedDataAsync` and `setSelectedDataAsync` methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="0e016-p102">`getSelectedDataAsync` Метод работает только для выделенного пользователем фрагмента. Если необходимо сохранить выделенный фрагмент в документе, чтобы для чтения и записи выполнялся один и тот же выделенный фрагмент, необходимо добавить привязку с помощью метода [Bindings. addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (или создать привязку с одним из других методов "аддфром" объекта [Bindings](/javascript/api/office/office.bindings) ). Сведения о том, как создать привязку к области документа, а затем читать и записывать в привязку, можно в статье [Привязка к областям в документе или электронной таблице](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="0e016-p102">The `getSelectedDataAsync` method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](/javascript/api/office/office.bindings) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="0e016-109">Чтение выбранных данных</span><span class="sxs-lookup"><span data-stu-id="0e016-109">Read selected data</span></span>


<span data-ttu-id="0e016-110">В примере ниже показано, как получить данные из выделенного фрагмента в документе с помощью метода [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="0e016-110">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


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

<span data-ttu-id="0e016-p103">В этом примере первый параметр _coercionType_ задается как `Office.CoercionType.Text` (также можно указать этот параметр с помощью литеральной строки `"text"`). Это означает, что свойство [value](/javascript/api/office/office.asyncresult#status) объекта [asyncResult](/javascript/api/office/office.asyncresult) , доступное из параметра _asyncResult_ в функции обратного вызова, будет возвращать **строку** , содержащую выделенный текст в документе. Задание различных типов приведения приведет к различным значениям. [Office. CoercionType](/javascript/api/office/office.coerciontype) — это перечисление доступных значений типов приведения. `Office.CoercionType.Text` возвращает строку "Text".</span><span class="sxs-lookup"><span data-stu-id="0e016-p103">In this example, the first  _coercionType_ parameter is specified as `Office.CoercionType.Text` (you can also specify this parameter by using the literal string `"text"`). This means that the [value](/javascript/api/office/office.asyncresult#status) property of the [AsyncResult](/javascript/api/office/office.asyncresult) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](/javascript/api/office/office.coerciontype) is an enumeration of available coercion type values. `Office.CoercionType.Text` evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="0e016-p104">**В каких случаях следует использовать матрицу и таблицу coercionType для доступа к данным?** Если вы хотите, чтобы выбранные табличные данные динамически увеличивались при добавлении строк и столбцов, и необходимо работать с заголовками таблиц, следует использовать тип данных table (указав параметр _coercionType_ `getSelectedDataAsync` метода AS `"table"` или `Office.CoercionType.Table`). Добавление строк и столбцов в структуре данных поддерживается как в табличных, так и в матричных данных, но добавление строк и столбцов поддерживается только для табличных данных. Если вы не планируете добавлять строки и столбцы, а ваши данные не нуждаются в функциях заголовков, следует использовать тип данных Matrix (указав параметр _coercionType_ `getSelectedDataAsync` метода AS `"matrix"` или `Office.CoercionType.Matrix`), что обеспечивает более простую модель взаимодействия с данными.</span><span class="sxs-lookup"><span data-stu-id="0e016-p104">**When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the `getSelectedDataAsync` method as `"table"` or `Office.CoercionType.Table`). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of `getSelectedDataAsync` method as `"matrix"` or `Office.CoercionType.Matrix`), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="0e016-p105">Анонимная функция, передаваемая функции в качестве второго параметра _обратного вызова_ , выполняется по завершении `getSelectedDataAsync` операции. Функция вызывается с одним параметром _asyncResult_, который содержит результат и состояние вызова. Если происходит сбой вызова, свойство [Error](/javascript/api/office/office.asyncresult#asynccontext) `AsyncResult` объекта предоставляет доступ к объекту [Error](/javascript/api/office/office.error) . Вы можете проверить значение свойства [Error.Name](/javascript/api/office/office.error#name) и [Error. Message](/javascript/api/office/office.error#message) , чтобы определить, почему не удалось выполнить операцию SET. В противном случае отображается выделенный текст в документе.</span><span class="sxs-lookup"><span data-stu-id="0e016-p105">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the `getSelectedDataAsync` operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](/javascript/api/office/office.asyncresult#asynccontext) property of the `AsyncResult` object provides access to the [Error](/javascript/api/office/office.error) object. You can check the value of the [Error.name](/javascript/api/office/office.error#name) and [Error.message](/javascript/api/office/office.error#message) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="0e016-p106">Свойство [asyncResult. status](/javascript/api/office/office.asyncresult#error) используется в операторе **If** , чтобы проверить, успешно ли выполнен вызов. [Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#status) — это перечисление `AsyncResult.status` доступных значений свойств. `Office.AsyncResultStatus.Failed` оценивается как строка "Failed" (а также может быть указана как строка этого литерала).</span><span class="sxs-lookup"><span data-stu-id="0e016-p106">The [AsyncResult.status](/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) is an enumeration of available `AsyncResult.status` property values. `Office.AsyncResultStatus.Failed` evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="0e016-128">Запись данных в выделение</span><span class="sxs-lookup"><span data-stu-id="0e016-128">Write data to the selection</span></span>


<span data-ttu-id="0e016-129">В следующем примере показано, как записать в выделение строку "Hello World!".</span><span class="sxs-lookup"><span data-stu-id="0e016-129">The following example shows how to set the selection to show "Hello World!".</span></span>


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

<span data-ttu-id="0e016-p107">Передача в параметре _data_ других типов объектов может привести к разным результатам. Результат зависит от текущего выделения в документе, от ведущего приложения, а также от возможности приведения переданных данных применительно к текущему выделению.</span><span class="sxs-lookup"><span data-stu-id="0e016-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="0e016-p108">Анонимная функция, передаваемая в метод [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) в качестве параметра _callback_ , выполняется при завершении асинхронного вызова. При записи данных в выделенный фрагмент с помощью `setSelectedDataAsync` метода параметр _asyncResult_ обратного вызова предоставляет доступ только к состоянию вызова и объекту [Error](/javascript/api/office/office.error) , если вызов завершается с ошибкой.</span><span class="sxs-lookup"><span data-stu-id="0e016-p108">The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the `setSelectedDataAsync` method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](/javascript/api/office/office.error) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="0e016-134">Начиная с выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel в Интернете, вы можете [задать форматирование при записи таблицы в текущую выделенную область](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="0e016-134">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel on the web, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="0e016-135">Обнаружение изменений в выделенной области</span><span class="sxs-lookup"><span data-stu-id="0e016-135">Detect changes in the selection</span></span>


<span data-ttu-id="0e016-136">В примере ниже показано, как определять изменения в выделенном фрагменте, используя метод [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) для добавления обработчика события [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) в документе.</span><span class="sxs-lookup"><span data-stu-id="0e016-136">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event on the document.</span></span>


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

<span data-ttu-id="0e016-p109">Первый параметр _EventType_ указывает имя события, на которое необходимо подписаться. Передача строки `"documentSelectionChanged"` для этого параметра эквивалентна передаче типа `Office.EventType.DocumentSelectionChanged` события для перечисления [Office. EventType](/javascript/api/office/office.eventtype) .</span><span class="sxs-lookup"><span data-stu-id="0e016-p109">The first  _eventType_ parameter specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the `Office.EventType.DocumentSelectionChanged` event type of the [Office.EventType](/javascript/api/office/office.eventtype) enumeration.</span></span>

<span data-ttu-id="0e016-p110">Анонимная функция `myHander()`, передаваемая в эту функцию в качестве второго параметра _handler_, представляет собой обработчик событий, который выполняется при изменении выделенного фрагмента в документе. При вызове этой функции передается единственный параметр _eventArgs_, который после завершения асинхронной операции будет содержать ссылку на объект [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs). Вы можете использовать свойство [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) для доступа к документу, создавшему событие.</span><span class="sxs-lookup"><span data-stu-id="0e016-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="0e016-p111">Можно добавить несколько обработчиков событий для данного события, повторно вызвав `addHandlerAsync` метод и передав дополнительную функцию обработчика событий для параметра _handler_ . Это будет работать правильно, если имя каждой функции обработчика событий уникально.</span><span class="sxs-lookup"><span data-stu-id="0e016-p111">You can add multiple event handlers for a given event by calling the `addHandlerAsync` method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="0e016-144">Отключение обнаружения изменений в выделенной области</span><span class="sxs-lookup"><span data-stu-id="0e016-144">Stop detecting changes in the selection</span></span>


<span data-ttu-id="0e016-145">В примере ниже показано, как остановить прослушивание события [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs), вызвав метод [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="0e016-145">The following example shows how to stop listening to the [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event by calling the [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="0e016-146">Имя `myHandler` функции, передаваемое в качестве второго параметра _handler_ , задает обработчик событий, который будет удален из `SelectionChanged` события.</span><span class="sxs-lookup"><span data-stu-id="0e016-146">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the `SelectionChanged` event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="0e016-147">Если необязательный параметр _handler_ опущен при вызове `removeHandlerAsync` метода, все обработчики событий для указанного объекта _EventType_ будут удалены.</span><span class="sxs-lookup"><span data-stu-id="0e016-147">If the optional  _handler_ parameter is omitted when the `removeHandlerAsync` method is called, all event handlers for the specified _eventType_ will be removed.</span></span>
