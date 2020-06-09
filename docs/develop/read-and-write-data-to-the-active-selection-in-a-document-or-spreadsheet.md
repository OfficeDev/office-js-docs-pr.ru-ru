---
title: Чтение и запись данных в текущую выделенную область документа или электронной таблицы
description: Сведения о том, как считывать и записывать данные в активный выбор в документе Word или в электронной таблице Excel.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: eb6c4d89e9c66ee3cda012c21601cb7454e73ae8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609396"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="1ebcf-103">Считывание и запись данных в активное выделение документа или таблицы</span><span class="sxs-lookup"><span data-stu-id="1ebcf-103">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="1ebcf-104">Объект [Document](/javascript/api/office/office.document) предоставляет методы, с помощью которых можно выполнять операции чтения и записи данных над текущим фрагментом, выделенным пользователем, в документе или электронной таблице.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-104">The [Document](/javascript/api/office/office.document) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet.</span></span> <span data-ttu-id="1ebcf-105">Для этого `Document` объект предоставляет `getSelectedDataAsync` `setSelectedDataAsync` методы и.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-105">To do that, the `Document` object provides the `getSelectedDataAsync` and `setSelectedDataAsync` methods.</span></span> <span data-ttu-id="1ebcf-106">Кроме того, в данной статье рассказывается, как считывать и записывать данные, а также создавать обработчики событий для обнаружения изменений в выделенном пользователем фрагменте.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-106">This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="1ebcf-107">`getSelectedDataAsync`Метод работает только для выделенного пользователем фрагмента.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-107">The `getSelectedDataAsync` method only works against the user's current selection.</span></span> <span data-ttu-id="1ebcf-108">Если необходимо сохранить выделенный фрагмент в документе, чтобы для чтения и записи выполнялся один и тот же выделенный фрагмент, необходимо добавить привязку с помощью метода [Bindings. addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (или создать привязку с одним из других методов "аддфром" объекта [Bindings](/javascript/api/office/office.bindings) ).</span><span class="sxs-lookup"><span data-stu-id="1ebcf-108">If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](/javascript/api/office/office.bindings) object).</span></span> <span data-ttu-id="1ebcf-109">Сведения о том, как создать привязку к области документа, а затем читать и записывать в привязку, можно в статье [Привязка к областям в документе или электронной таблице](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="1ebcf-109">For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="1ebcf-110">Чтение выбранных данных</span><span class="sxs-lookup"><span data-stu-id="1ebcf-110">Read selected data</span></span>


<span data-ttu-id="1ebcf-111">В примере ниже показано, как получить данные из выделенного фрагмента в документе с помощью метода [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="1ebcf-111">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


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

<span data-ttu-id="1ebcf-112">В этом примере первый параметр _coercionType_ задается как `Office.CoercionType.Text` (также можно указать этот параметр с помощью литеральной строки `"text"` ).</span><span class="sxs-lookup"><span data-stu-id="1ebcf-112">In this example, the first  _coercionType_ parameter is specified as `Office.CoercionType.Text` (you can also specify this parameter by using the literal string `"text"`).</span></span> <span data-ttu-id="1ebcf-113">Это означает, что свойство [value](/javascript/api/office/office.asyncresult#status) объекта [AsyncResult](/javascript/api/office/office.asyncresult), доступного из параметра _asyncResult_ в функции обратного вызова, возвратит **строку**, содержащую выделенный текст в документе.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-113">This means that the [value](/javascript/api/office/office.asyncresult#status) property of the [AsyncResult](/javascript/api/office/office.asyncresult) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document.</span></span> <span data-ttu-id="1ebcf-114">Если вы укажете какой-либо другой тип приведения, то получите другие значения.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-114">Specifying different coercion types will result in different values.</span></span> <span data-ttu-id="1ebcf-115">[Office.CoercionType](/javascript/api/office/office.coerciontype) — это перечисление значений доступных типов приведений.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-115">[Office.CoercionType](/javascript/api/office/office.coerciontype) is an enumeration of available coercion type values.</span></span> <span data-ttu-id="1ebcf-116">`Office.CoercionType.Text`Возвращает строку "Text".</span><span class="sxs-lookup"><span data-stu-id="1ebcf-116">`Office.CoercionType.Text` evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="1ebcf-117">**В каких случаях следует использовать для доступа к данным матрицы, а в каких — coercionType?**</span><span class="sxs-lookup"><span data-stu-id="1ebcf-117">**When should you use the matrix versus table coercionType for data access?**</span></span> <span data-ttu-id="1ebcf-118">Если вы хотите, чтобы выбранные табличные данные динамически увеличивались при добавлении строк и столбцов, и необходимо работать с заголовками таблиц, следует использовать тип данных table (указав параметр _coercionType_ `getSelectedDataAsync` метода AS `"table"` или `Office.CoercionType.Table` ).</span><span class="sxs-lookup"><span data-stu-id="1ebcf-118">If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the `getSelectedDataAsync` method as `"table"` or `Office.CoercionType.Table`).</span></span> <span data-ttu-id="1ebcf-119">Добавление строк и столбцов в структуре данных поддерживается как табличными, так и матричными данными, но присоединение строк и столбцов поддерживается только табличными данными.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-119">Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data.</span></span> <span data-ttu-id="1ebcf-120">Если вы не планируете добавлять строки и столбцы, а ваши данные не нуждаются в функциях заголовков, следует использовать тип данных Matrix (указав параметр _coercionType_ `getSelectedDataAsync` метода AS `"matrix"` или `Office.CoercionType.Matrix` ), что обеспечивает более простую модель взаимодействия с данными.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-120">If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of `getSelectedDataAsync` method as `"matrix"` or `Office.CoercionType.Matrix`), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="1ebcf-121">Анонимная функция, передаваемая функции в качестве второго параметра _обратного вызова_ , выполняется по `getSelectedDataAsync` завершении операции.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-121">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the `getSelectedDataAsync` operation is completed.</span></span> <span data-ttu-id="1ebcf-122">При вызове функции передается один параметр _asyncResult_, который содержит результат и сведения о состоянии вызова.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-122">The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call.</span></span> <span data-ttu-id="1ebcf-123">Если происходит сбой вызова, свойство [Error](/javascript/api/office/office.asyncresult#asynccontext) `AsyncResult` объекта предоставляет доступ к объекту [Error](/javascript/api/office/office.error) .</span><span class="sxs-lookup"><span data-stu-id="1ebcf-123">If the call fails, the [error](/javascript/api/office/office.asyncresult#asynccontext) property of the `AsyncResult` object provides access to the [Error](/javascript/api/office/office.error) object.</span></span> <span data-ttu-id="1ebcf-124">Вы можете проверить значение свойств [Error.name](/javascript/api/office/office.error#name) и [Error.message](/javascript/api/office/office.error#message), чтобы определить, почему операция завершилась с ошибкой.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-124">You can check the value of the [Error.name](/javascript/api/office/office.error#name) and [Error.message](/javascript/api/office/office.error#message) properties to determine why the set operation failed.</span></span> <span data-ttu-id="1ebcf-125">В противном случае будет отображен выделенный в документе текст.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-125">Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="1ebcf-126">Свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#error) используется в выражении **if** для проверки того, успешно ли выполнен вызов.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-126">The [AsyncResult.status](/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded.</span></span> <span data-ttu-id="1ebcf-127">[Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#status) — это перечисление доступных `AsyncResult.status` значений свойств.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-127">[Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) is an enumeration of available `AsyncResult.status` property values.</span></span> <span data-ttu-id="1ebcf-128">`Office.AsyncResultStatus.Failed`оценивается как строка "Failed" (а также может быть указана как строка этого литерала).</span><span class="sxs-lookup"><span data-stu-id="1ebcf-128">`Office.AsyncResultStatus.Failed` evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="1ebcf-129">Запись данных в выделение</span><span class="sxs-lookup"><span data-stu-id="1ebcf-129">Write data to the selection</span></span>


<span data-ttu-id="1ebcf-130">В следующем примере показано, как записать в выделение строку "Hello World!".</span><span class="sxs-lookup"><span data-stu-id="1ebcf-130">The following example shows how to set the selection to show "Hello World!".</span></span>


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

<span data-ttu-id="1ebcf-p107">Передача в параметре _data_ других типов объектов может привести к разным результатам. Результат зависит от текущего выделения в документе, от ведущего приложения, а также от возможности приведения переданных данных применительно к текущему выделению.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="1ebcf-133">Анонимная функция, которая передается в метод [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) в качестве параметра _callback_, выполняется после завершения асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-133">The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed.</span></span> <span data-ttu-id="1ebcf-134">При записи данных в выделенный фрагмент с помощью `setSelectedDataAsync` метода параметр _asyncResult_ обратного вызова предоставляет доступ только к состоянию вызова и объекту [Error](/javascript/api/office/office.error) , если вызов завершается с ошибкой.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-134">When you write data to the selection by using the `setSelectedDataAsync` method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](/javascript/api/office/office.error) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="1ebcf-135">Начиная с выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel в Интернете, вы можете [задать форматирование при записи таблицы в текущую выделенную область](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="1ebcf-135">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel on the web, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="1ebcf-136">Обнаружение изменений в выделенной области</span><span class="sxs-lookup"><span data-stu-id="1ebcf-136">Detect changes in the selection</span></span>


<span data-ttu-id="1ebcf-137">В примере ниже показано, как определять изменения в выделенном фрагменте, используя метод [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) для добавления обработчика события [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) в документе.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-137">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event on the document.</span></span>


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

<span data-ttu-id="1ebcf-138">Первый параметр _eventType_ задает имя события для подписки.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-138">The first  _eventType_ parameter specifies the name of the event to subscribe to.</span></span> <span data-ttu-id="1ebcf-139">Передача строки `"documentSelectionChanged"` для этого параметра эквивалентна передаче `Office.EventType.DocumentSelectionChanged` типа события для перечисления [Office. EventType](/javascript/api/office/office.eventtype) .</span><span class="sxs-lookup"><span data-stu-id="1ebcf-139">Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the `Office.EventType.DocumentSelectionChanged` event type of the [Office.EventType](/javascript/api/office/office.eventtype) enumeration.</span></span>

<span data-ttu-id="1ebcf-p110">Анонимная функция `myHander()`, передаваемая в эту функцию в качестве второго параметра _handler_, представляет собой обработчик событий, который выполняется при изменении выделенного фрагмента в документе. При вызове этой функции передается единственный параметр _eventArgs_, который после завершения асинхронной операции будет содержать ссылку на объект [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs). Вы можете использовать свойство [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) для доступа к документу, создавшему событие.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="1ebcf-143">Можно добавить несколько обработчиков событий для данного события, повторно вызвав `addHandlerAsync` метод и передав дополнительную функцию обработчика событий для параметра _handler_ .</span><span class="sxs-lookup"><span data-stu-id="1ebcf-143">You can add multiple event handlers for a given event by calling the `addHandlerAsync` method again and passing in an additional event handler function for the _handler_ parameter.</span></span> <span data-ttu-id="1ebcf-144">Это будет работать правильно, поскольку имя каждой функции обработчика событий уникально.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-144">This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="1ebcf-145">Отключение обнаружения изменений в выделенной области</span><span class="sxs-lookup"><span data-stu-id="1ebcf-145">Stop detecting changes in the selection</span></span>


<span data-ttu-id="1ebcf-146">В примере ниже показано, как остановить прослушивание события [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs), вызвав метод [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="1ebcf-146">The following example shows how to stop listening to the [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event by calling the [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="1ebcf-147">`myHandler`Имя функции, передаваемое в качестве второго параметра _handler_ , задает обработчик событий, который будет удален из `SelectionChanged` события.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-147">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the `SelectionChanged` event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="1ebcf-148">Если необязательный параметр _handler_ опущен при `removeHandlerAsync` вызове метода, все обработчики событий для указанного объекта _EventType_ будут удалены.</span><span class="sxs-lookup"><span data-stu-id="1ebcf-148">If the optional  _handler_ parameter is omitted when the `removeHandlerAsync` method is called, all event handlers for the specified _eventType_ will be removed.</span></span>
