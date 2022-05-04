---
title: Чтение и запись данных в текущую выделенную область документа или электронной таблицы
description: Узнайте, как читать и записывать данные в активное выделение в документе Word или Excel электронной таблице.
ms.date: 01/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 360701bc43a7fc63f8447ff9a068256d187e2a70
ms.sourcegitcommit: 5bf28c447c5b60e2cc7e7a2155db66cd9fe2ab6b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/04/2022
ms.locfileid: "65187317"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Считывание и запись данных в активное выделение документа или таблицы

Объект [Document](/javascript/api/office/office.document) предоставляет методы, с помощью которых можно выполнять операции чтения и записи данных над текущим фрагментом, выделенным пользователем, в документе или электронной таблице. Для этого объект предоставляет `Document` методы `getSelectedDataAsync` и методы `setSelectedDataAsync` . Кроме того, в данной статье рассказывается, как считывать и записывать данные, а также создавать обработчики событий для обнаружения изменений в выделенном пользователем фрагменте.

Этот `getSelectedDataAsync` метод работает только с текущим выбором пользователя. Если необходимо сохранить выделение в документе, чтобы один и тот же фрагмент был доступен для чтения и записи в сеансах запуска надстройки, необходимо добавить привязку с помощью метода [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) (или создать привязку с помощью одного из других методов addFrom объекта [Bindings](/javascript/api/office/office.bindings) ). Сведения о создании привязки к области документа, а затем чтении и записи в привязку см. в разделе "Привязка к регионам" документа или [электронной таблицы](bind-to-regions-in-a-document-or-spreadsheet.md).

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

В этом примере указывается первый параметр _coercionType_`Office.CoercionType.Text` (этот параметр также можно указать с помощью строки литерала`"text"`). Это означает, что свойство [value](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) объекта [AsyncResult](/javascript/api/office/office.asyncresult), доступного из параметра _asyncResult_ в функции обратного вызова, возвратит **строку**, содержащую выделенный текст в документе. Если вы укажете какой-либо другой тип приведения, то получите другие значения. [Office.CoercionType](/javascript/api/office/office.coerciontype) — это перечисление значений доступных типов приведений. `Office.CoercionType.Text` вычисляет строку "text".

> [!TIP]
> **В каких случаях следует использовать для доступа к данным матрицы, а в каких — coercionType?** Если при добавлении строк и столбцов необходимо динамически увеличивать выбранные табличные данные и работать с заголовками таблиц, следует использовать табличный тип данных (указав параметр _coercionType_ `getSelectedDataAsync` `"table"` `Office.CoercionType.Table`метода как или). Добавление строк и столбцов в структуре данных поддерживается как табличными, так и матричными данными, но присоединение строк и столбцов поддерживается только табличными данными. Если вы не планируете добавлять строки и столбцы и данные не требуют функциональных возможностей заголовков, следует использовать тип данных матрицы (указав параметр  _coercionType_ `getSelectedDataAsync` `"matrix"` `Office.CoercionType.Matrix`метода как или), который обеспечивает более простую модель взаимодействия с данными.

Анонимная функция, которая передается в функцию в качестве второго _параметра,_ обратного вызова, выполняется после `getSelectedDataAsync` завершения операции. При вызове функции передается один параметр _asyncResult_, который содержит результат и сведения о состоянии вызова. Если вызов завершается сбоем, свойство [ошибки](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) `AsyncResult` объекта предоставляет доступ к [объекту Error](/javascript/api/office/office.error) . Вы можете проверить значение свойств [Error.name](/javascript/api/office/office.error#office-office-error-name-member) и [Error.message](/javascript/api/office/office.error#office-office-error-message-member), чтобы определить, почему операция завершилась с ошибкой. В противном случае будет отображен выделенный в документе текст.

Свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) используется в выражении **if** для проверки того, успешно ли выполнен вызов. [Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) — это перечисление доступных значений `AsyncResult.status` свойств. `Office.AsyncResultStatus.Failed` вычисляет строку "failed" (и, опять же, можно указать в качестве этой строки литерала).

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

Передача в параметре _data_ других типов объектов может привести к разным результатам. Результат зависит от того, что в настоящее время выбрано в документе, какое Office клиентское приложение размещает надстройку, и можно ли привести переданные данные к текущему выбору.

Анонимная функция, которая передается в метод [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) в качестве параметра _callback_, выполняется после завершения асинхронного вызова. При записи `setSelectedDataAsync` данных в выборку с помощью метода параметр _asyncResult_ обратного вызова предоставляет доступ только к статусу вызова и к объекту [Error](/javascript/api/office/office.error) в случае сбоя вызова.

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

Первый параметр _, eventType_, указывает имя события, на которое необходимо подписаться. Передача строки `"documentSelectionChanged"` для этого параметра эквивалентна `Office.EventType.DocumentSelectionChanged` передаче типа события Office[. Перечисление EventType](/javascript/api/office/office.eventtype).

Функция  `myHandler()` , которая передается в функцию в качестве второго параметра _,_ обработчика, является обработчиком событий, который выполняется при изменении выделения в документе. При вызове этой функции передается единственный параметр _eventArgs_, который после завершения асинхронной операции будет содержать ссылку на объект [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs). Вы можете использовать свойство [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#office-office-documentselectionchangedeventargs-document-member) для доступа к документу, создавшему событие.

> [!NOTE]
> Вы можете добавить несколько обработчиков `addHandlerAsync` событий для данного события, вызвав метод еще раз и передав дополнительную функцию обработчика событий для _параметра обработчика_ . Это будет работать правильно, поскольку имя каждой функции обработчика событий уникально.

## <a name="stop-detecting-changes-in-the-selection"></a>Отключение обнаружения изменений в выделенной области

В примере ниже показано, как остановить прослушивание события [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs), вызвав метод [document.removeHandlerAsync](/javascript/api/office/office.document#office-office-document-removehandlerasync-member(1)).

```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Имя  `myHandler` функции, которое передается в качестве второго параметра, _обработчика_, указывает обработчик событий, который будет удален из `SelectionChanged` события.

> [!IMPORTANT]
> Если  _необязательный_ `removeHandlerAsync` параметр обработчика опущен при вызове метода, все обработчики событий для указанного _eventType_ будут удалены.
