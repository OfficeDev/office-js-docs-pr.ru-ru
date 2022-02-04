---
title: Чтение и запись данных в текущую выделенную область документа или электронной таблицы
description: 'Узнайте, как читать и записывать данные для активного выбора в документе Word или Excel таблице.'
ms.date: 01/31/2022
ms.localizationpriority: medium
---


# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Считывание и запись данных в активное выделение документа или таблицы

Объект [Document](/javascript/api/office/office.document) предоставляет методы, с помощью которых можно выполнять операции чтения и записи данных над текущим фрагментом, выделенным пользователем, в документе или электронной таблице. Для этого объект предоставляет `Document` средства и `getSelectedDataAsync` методы `setSelectedDataAsync` . Кроме того, в данной статье рассказывается, как считывать и записывать данные, а также создавать обработчики событий для обнаружения изменений в выделенном пользователем фрагменте.

Метод `getSelectedDataAsync` работает только в отношении текущего выбора пользователя. Если необходимо сохранить выбор в документе, чтобы один и тот же выбор был доступен для чтения и записи во всех сеансах запуска надстройки, необходимо добавить привязку с помощью метода [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) (или создать привязку с одним из других методов addFrom объекта [Bindings](/javascript/api/office/office.bindings) ). Сведения о создании привязки к региону документа, а затем чтении и записи к привязке см. в книге [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Чтение выбранных данных


В примере ниже показано, как получить данные из выделенного фрагмента в документе с помощью метода [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)).


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

В этом примере указывается первый параметр  _coercionType_ `Office.CoercionType.Text` (можно также указать этот параметр с помощью литеральной строки `"text"`). Это означает, что свойство [value](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) объекта [AsyncResult](/javascript/api/office/office.asyncresult), доступного из параметра _asyncResult_ в функции обратного вызова, возвратит **строку**, содержащую выделенный текст в документе. Если вы укажете какой-либо другой тип приведения, то получите другие значения. [Office.CoercionType](/javascript/api/office/office.coerciontype) — это перечисление значений доступных типов приведений. `Office.CoercionType.Text` оценивает строку "text".


> [!TIP]
> **В каких случаях следует использовать для доступа к данным матрицы, а в каких — coercionType?** Если для динамического роста выбранных табулярных данных при добавлении строк и столбцов необходимо работать с загонами таблиц, следует использовать тип данных таблицы (укажите параметр _coercionType_ `getSelectedDataAsync` `"table"` метода как или `Office.CoercionType.Table`). Добавление строк и столбцов в структуре данных поддерживается как табличными, так и матричными данными, но присоединение строк и столбцов поддерживается только табличными данными. Если вы не планируете добавлять строки и столбцы, а ваши данные не требуют функции загона, то следует использовать тип матричных данных (указав параметр  _coercionType_ `getSelectedDataAsync` `"matrix"` `Office.CoercionType.Matrix`метода как или), который обеспечивает более простую модель взаимодействия с данными.

Анонимная функция, которая передается в функцию в  качестве второго параметра обратного вызова, выполняется по завершению `getSelectedDataAsync` операции. При вызове функции передается один параметр _asyncResult_, который содержит результат и сведения о состоянии вызова. Если вызов не удается, свойство [ошибки](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) `AsyncResult` объекта предоставляет доступ к [объекту Error](/javascript/api/office/office.error) . Вы можете проверить значение свойств [Error.name](/javascript/api/office/office.error#office-office-error-name-member) и [Error.message](/javascript/api/office/office.error#office-office-error-message-member), чтобы определить, почему операция завершилась с ошибкой. В противном случае будет отображен выделенный в документе текст.

Свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) используется в выражении **if** для проверки того, успешно ли выполнен вызов. [Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) — это переумежение доступных `AsyncResult.status` значений свойств. `Office.AsyncResultStatus.Failed` оценивает строку "не удалось" (и, опять же, также может быть указан в качестве этой буквальной строки).


## <a name="write-data-to-the-selection"></a>Запись данных в выделение


В следующем примере показано, как записать в выделение строку "Hello World!".


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

Передача в параметре _data_ других типов объектов может привести к разным результатам. Результат зависит от того, что в настоящее время выбрано в документе, в котором Office клиентского приложения размещена ваша надстройка, и можно ли принудить переданные данные к текущему выбору.

Анонимная функция, которая передается в метод [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) в качестве параметра _callback_, выполняется после завершения асинхронного вызова. При записи `setSelectedDataAsync` данных в выбор с помощью метода параметр _asyncResult_ от вызываемого вызова предоставляет доступ только к статусу вызова и к объекту [Error](/javascript/api/office/office.error) , если вызов не удается.

> [!NOTE]
> Начиная с выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel в Интернете, вы можете [задать форматирование при записи таблицы в текущую выделенную область](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Обнаружение изменений в выделенной области


В примере ниже показано, как определять изменения в выделенном фрагменте, используя метод [Document.addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) для добавления обработчика события [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) в документе.


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

Первый параметр _eventType_ задает имя события для подписки. Передача строки `"documentSelectionChanged"` для этого параметра эквивалентна `Office.EventType.DocumentSelectionChanged` передаче типа [события Office. Переумерия EventType](/javascript/api/office/office.eventtype).

Анонимная функция `myHandler()`, передаваемая в эту функцию в качестве второго параметра _handler_, представляет собой обработчик событий, который выполняется при изменении выделенного фрагмента в документе. При вызове этой функции передается единственный параметр _eventArgs_, который после завершения асинхронной операции будет содержать ссылку на объект [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs). Вы можете использовать свойство [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#office-office-documentselectionchangedeventargs-document-member) для доступа к документу, создавшему событие.


> [!NOTE]
> Вы можете добавить несколько обработчиков `addHandlerAsync` событий для данного события, снова позвонив методу и передав дополнительную функцию обработчику событий для параметра _обработчик_ . Это будет работать правильно, поскольку имя каждой функции обработчика событий уникально.


## <a name="stop-detecting-changes-in-the-selection"></a>Отключение обнаружения изменений в выделенной области


В примере ниже показано, как остановить прослушивание события [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs), вызвав метод [document.removeHandlerAsync](/javascript/api/office/office.document#office-office-document-removehandlerasync-member(1)).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Имя  `myHandler` функции, которое передается в качестве параметра _второго_ обработера, указывает обработник событий, который будет удален из `SelectionChanged` события.


> [!IMPORTANT]
> Если  _параметр_ необязательный `removeHandlerAsync` обработчик будет опущен, когда метод называется, все обработчики событий для указанного _eventType_ будут удалены.
