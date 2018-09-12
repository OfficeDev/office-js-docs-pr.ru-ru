---
title: Чтение и запись данных в текущую выделенную область документа или электронной таблицы
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d1c8fcdeec8d92fd3f77e169dc24715f7c5e9964
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944988"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Чтение и запись данных в текущую выделенную область документа или электронной таблицы

Объект [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) предоставляет методы, с помощью которых можно выполнять операции чтения и записи данных над текущим фрагментом, выделенным пользователем, в документе или электронной таблице. Для этого в объекте **Document** имеются методы **getSelectedDataAsync** и **setSelectedDataAsync**. Кроме того, в данной статье рассказывается, как считывать и записывать данные, а также создавать обработчики событий для обнаружения изменений в выделенном пользователем фрагменте.

Метод **getSelectedDataAsync** работает только для текущего фрагмента, выделенного пользователем. Если вам необходимо сохранить выделенный фрагмент в документе, чтобы он был доступен для чтения и записи во время последующих сеансов работы надстройки, необходимо добавить привязку с помощью метода [Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) (или создать привязку с помощью любого метода addFrom объекта [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js)). Дополнительные сведения о том, как создать привязку к области в документе, а также о чтении и записи данных через привязку см. в разделе [Привязка к областям в документе или электронной таблице](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Чтение выбранных данных


В примере ниже показано, как получить данные из выделенного фрагмента в документе с помощью метода [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-).


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

В этом примере первый параметр _coercionType_ имеет значение **Office.CoercionType.Text** (вы также можете указать этот параметр, используя строку литерала `"text"`). Это означает, что свойство [value](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) объекта [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js), доступного из параметра _asyncResult_ в функции обратного вызова, возвратит **строку**, содержащую выделенный текст в документе. Если вы укажете какой-либо другой тип приведения, то получите другие значения. [Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) — это перечисление значений доступных типов приведений. **Office.CoercionType.Text** имеет значение text.


> [!TIP]
> **Информация о выборе матричного или табличного coercionType для доступа к данным.** Если предполагается динамический рост данных таблицы с добавлением строк и столбцов, при этом требуется обрабатывать заголовки таблицы, следует использовать табличные данные (указав параметр _coercionType_ метода **getSelectedDataAsync** в виде `"table"` или **Office.CoercionType.Table**). Добавление строк и столбцов в структуре данных поддерживается как табличными, так и матричными данными, но присоединение строк и столбцов поддерживается только табличными данными. В отсутствие необходимости добавления строк и столбцов, если при этом не требуется использовать заголовки, следует выбрать матричные данные (указав параметр _coercionType_ метода **getSelecteDataAsync** в виде `"matrix"` или **Office.CoercionType.Matrix**), что позволяет использовать упрощенный способ взаимодействия с данными.

Анонимная функция, которая передается в функцию в качестве второго параметра _callback_, выполняется после завершения операции **getSelectedDataAsync**. При вызове функции передается один параметр _asyncResult_, который содержит результат и сведения о состоянии вызова. Если вызов завершается с ошибкой, свойство [error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#asynccontext) объекта **AsyncResult** предоставляет доступ к объекту [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js). Вы можете проверить значение свойств [Error.name](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#name) и [Error.message](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#message), чтобы определить, почему операция завершилась с ошибкой. В противном случае будет отображен выделенный в документе текст.

Свойство [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#error) используется в выражении **if** для проверки того, успешно ли выполнен вызов. [Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) — это перечисление доступных значений свойства **AsyncResult.status**. **Office.AsyncResultStatus.Failed** имеет значение failed (и его можно указать в виде строки литералов).


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

Передача в параметре _data_ других типов объектов может привести к разным результатам. Результат зависит от текущего выделения в документе, от ведущего приложения, а также от возможности приведения переданных данных применительно к текущему выделению.

Анонимная функция, которая передается в метод [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) в качестве параметра _callback_, выполняется после завершения асинхронного вызова. При записи данных в выделенный фрагмент с помощью метода **setSelectedDataAsync** параметр _asyncResult_ обратного вызова предоставляет доступ только к сведениям о состоянии вызова и к объекту [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) в случае сбоя вызова.

> [!NOTE]
> Начиная с выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel Online вы можете [задать форматирование при записи таблицы в текущую выделенную область](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Обнаружение изменений в выделенной области


В примере ниже показано, как определять изменения в выделенном фрагменте, используя метод [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) для добавления обработчика события [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) в документе.


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

Первый параметр _eventType_ задает имя события для подписки. Передача строки `"documentSelectionChanged"` для этого параметра эквивалентна передаче типа события **Office.EventType.DocumentSelectionChanged** перечисления [Office.EventType](https://docs.microsoft.com/javascript/api/office/office.eventtype?view=office-js).

Анонимная функция `myHander()`, передаваемая в эту функцию в качестве второго параметра _handler_, представляет собой обработчик событий, который выполняется при изменении выделенного фрагмента в документе. При вызове этой функции передается единственный параметр _eventArgs_, который после завершения асинхронной операции будет содержать ссылку на объект [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js). Вы можете использовать свойство [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js#document) для доступа к документу, создавшему событие.


> [!NOTE]
> Вы можете добавить несколько обработчиков событий для данного события, снова вызвав метод **addHandlerAsync** и передав дополнительную функцию обработчика события для параметра _handler_. Это будет работать правильно, поскольку имя каждой функции обработчика событий уникально.


## <a name="stop-detecting-changes-in-the-selection"></a>Отключение обнаружения изменений в выделенной области


В примере ниже показано, как остановить прослушивание события [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js), вызвав метод [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#removehandlerasync-eventtype--options--callback-).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Имя функции `myHandler`, передаваемое в качестве второго параметра _handler_, задает обработчик событий, который будет удален из события **SelectionChanged**.


> [!IMPORTANT]
> Если необязательный параметр _handler_ при вызове метода **removeHandlerAsync** не указывается, то все обработчики событий для указанного объекта _eventType_ будут удалены.

