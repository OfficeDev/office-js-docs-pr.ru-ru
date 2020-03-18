---
title: Чтение и запись данных в текущую выделенную область документа или электронной таблицы
description: Сведения о том, как считывать и записывать данные в активный выбор в документе Word или в электронной таблице Excel.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 83f3de5c522436ac06a0238781ee71de676297a1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718883"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Считывание и запись данных в активное выделение документа или таблицы

Объект [Document](/javascript/api/office/office.document) предоставляет методы, с помощью которых можно выполнять операции чтения и записи данных над текущим фрагментом, выделенным пользователем, в документе или электронной таблице. Для этого `Document` объект предоставляет методы `getSelectedDataAsync` и. `setSelectedDataAsync` Кроме того, в данной статье рассказывается, как считывать и записывать данные, а также создавать обработчики событий для обнаружения изменений в выделенном пользователем фрагменте.

`getSelectedDataAsync` Метод работает только для выделенного пользователем фрагмента. Если необходимо сохранить выделенный фрагмент в документе, чтобы для чтения и записи выполнялся один и тот же выделенный фрагмент, необходимо добавить привязку с помощью метода [Bindings. addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (или создать привязку с одним из других методов "аддфром" объекта [Bindings](/javascript/api/office/office.bindings) ). Сведения о том, как создать привязку к области документа, а затем читать и записывать в привязку, можно в статье [Привязка к областям в документе или электронной таблице](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Чтение выбранных данных


В примере ниже показано, как получить данные из выделенного фрагмента в документе с помощью метода [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-).


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

В этом примере первый параметр _coercionType_ задается как `Office.CoercionType.Text` (также можно указать этот параметр с помощью литеральной строки `"text"`). Это означает, что свойство [value](/javascript/api/office/office.asyncresult#status) объекта [AsyncResult](/javascript/api/office/office.asyncresult), доступного из параметра _asyncResult_ в функции обратного вызова, возвратит **строку**, содержащую выделенный текст в документе. Если вы укажете какой-либо другой тип приведения, то получите другие значения. [Office.CoercionType](/javascript/api/office/office.coerciontype) — это перечисление значений доступных типов приведений. `Office.CoercionType.Text`Возвращает строку "Text".


> [!TIP]
> **В каких случаях следует использовать для доступа к данным матрицы, а в каких — coercionType?** Если вы хотите, чтобы выбранные табличные данные динамически увеличивались при добавлении строк и столбцов, и необходимо работать с заголовками таблиц, следует использовать тип данных table (указав параметр _coercionType_ `getSelectedDataAsync` метода AS `"table"` или `Office.CoercionType.Table`). Добавление строк и столбцов в структуре данных поддерживается как табличными, так и матричными данными, но присоединение строк и столбцов поддерживается только табличными данными. Если вы не планируете добавлять строки и столбцы, а ваши данные не нуждаются в функциях заголовков, следует использовать тип данных Matrix (указав параметр _coercionType_ `getSelectedDataAsync` метода AS `"matrix"` или `Office.CoercionType.Matrix`), что обеспечивает более простую модель взаимодействия с данными.

Анонимная функция, передаваемая функции в качестве второго параметра _обратного вызова_ , выполняется по завершении `getSelectedDataAsync` операции. При вызове функции передается один параметр _asyncResult_, который содержит результат и сведения о состоянии вызова. Если происходит сбой вызова, свойство [Error](/javascript/api/office/office.asyncresult#asynccontext) `AsyncResult` объекта предоставляет доступ к объекту [Error](/javascript/api/office/office.error) . Вы можете проверить значение свойств [Error.name](/javascript/api/office/office.error#name) и [Error.message](/javascript/api/office/office.error#message), чтобы определить, почему операция завершилась с ошибкой. В противном случае будет отображен выделенный в документе текст.

Свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#error) используется в выражении **if** для проверки того, успешно ли выполнен вызов. [Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#status) — это перечисление `AsyncResult.status` доступных значений свойств. `Office.AsyncResultStatus.Failed`оценивается как строка "Failed" (а также может быть указана как строка этого литерала).


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

Анонимная функция, которая передается в метод [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) в качестве параметра _callback_, выполняется после завершения асинхронного вызова. При записи данных в выделенный фрагмент с помощью `setSelectedDataAsync` метода параметр _asyncResult_ обратного вызова предоставляет доступ только к состоянию вызова и объекту [Error](/javascript/api/office/office.error) , если вызов завершается с ошибкой.

> [!NOTE]
> Начиная с выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel в Интернете, вы можете [задать форматирование при записи таблицы в текущую выделенную область](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Обнаружение изменений в выделенной области


В примере ниже показано, как определять изменения в выделенном фрагменте, используя метод [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) для добавления обработчика события [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) в документе.


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

Первый параметр _eventType_ задает имя события для подписки. Передача строки `"documentSelectionChanged"` для этого параметра эквивалентна передаче типа `Office.EventType.DocumentSelectionChanged` события для перечисления [Office. EventType](/javascript/api/office/office.eventtype) .

Анонимная функция `myHander()`, передаваемая в эту функцию в качестве второго параметра _handler_, представляет собой обработчик событий, который выполняется при изменении выделенного фрагмента в документе. При вызове этой функции передается единственный параметр _eventArgs_, который после завершения асинхронной операции будет содержать ссылку на объект [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs). Вы можете использовать свойство [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) для доступа к документу, создавшему событие.


> [!NOTE]
> Можно добавить несколько обработчиков событий для данного события, повторно вызвав `addHandlerAsync` метод и передав дополнительную функцию обработчика событий для параметра _handler_ . Это будет работать правильно, поскольку имя каждой функции обработчика событий уникально.


## <a name="stop-detecting-changes-in-the-selection"></a>Отключение обнаружения изменений в выделенной области


В примере ниже показано, как остановить прослушивание события [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs), вызвав метод [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Имя `myHandler` функции, передаваемое в качестве второго параметра _handler_ , задает обработчик событий, который будет удален из `SelectionChanged` события.


> [!IMPORTANT]
> Если необязательный параметр _handler_ опущен при вызове `removeHandlerAsync` метода, все обработчики событий для указанного объекта _EventType_ будут удалены.
