---
title: Привязка к областям в документе или электронной таблице
description: Узнайте, как использовать привязку для обеспечения согласованного доступа к определенной области или элементу документа или электронной таблицы через идентификатор.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: b1624624a5664444f811b20d405fb6aefbd23d91
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958855"
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a>Привязка к областям в документе или электронной таблице

Благодаря доступу к данным на основе привязок контентные надстройки и надстройки области задач могут согласованно получать доступ к определенной области документа или электронной таблицы с помощью идентификатора. Прежде всего надстройке необходимо создать привязку, вызвав один из методов, сопоставляющих часть документа с уникальным идентификатором: [addFromPromptAsync], [addFromSelectionAsync] или [addFromNamedItemAsync]. После настройки привязки надстройка может использовать предоставленный идентификатор для доступа к данным, содержащимся в связанном регионе документа или электронной таблицы. Создание привязок предоставляет надстройке следующее значение.

- Разрешает доступ к общим структурам данных в поддерживаемых приложениях Office, таким как: таблицы, диапазоны или текст (связанная последовательность знаков).
- Позволяет производить операции чтения или записи без необходимости выделения пользователем фрагмента.
- Устанавливает отношение между надстройкой и данными в документе. Привязки сохраняются в документе и могут использоваться позже.

Установка привязки также позволяет подписываться на данные и выбирать изменения событий, относящиеся к конкретной области документа или электронной таблицы. Это означает, что надстройка уведомляется только об изменениях, происходящих внутри данной конкретной области, в отличие от изменений, затрагивающих в целом весь документ или электронную таблицу.

Объект [Bindings] предоставляет метод [getAllAsync], который обеспечивает доступ к набору всех привязок, установленных в этом документе или листе. Доступ к отдельной привязке можно получить по ее идентификатору с помощью привязки. [Метод getByIdAsync] или [функция Office.select] . Можно создать новые привязки, а также удалить существующие, используя один из перечисленных ниже методов объекта [Bindings]: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync] или [releaseByIdAsync].

## <a name="binding-types"></a>Типы привязок

Существует три различных типа привязок [][Office.BindingType], задаваемых с помощью параметра _bindingType_ при создании привязки с помощью методов [addFromSelectionAsync], [addFromPromptAsync] или [addFromNamedItemAsync].

1. **[Текстовая привязка][TextBinding]**. Выполняет привязку к области документа, которая может быть представлена как текст.

    В Word поддерживается большинство связанных выделений, тогда как в Excel для привязки текста можно использовать только выделения отдельных ячеек. Excel поддерживает только обычный текст, а Word — три формата: обычный текст, HTML и Open XML для Office.

1. **[Matrix Binding][MatrixBinding]** — привязка к фиксированной области документа, содержащего табличные данные без заголовков. Данные в матричной привязке записываются или считывается как двумерный **массив, который** в JavaScript реализуется как массив массивов. Например, две строки значений **string** в двух столбцах можно записать или прочитать как `[['a', 'b'], ['c', 'd']]`, а один столбец, состоящий из трех строк, — как `[['a'], ['b'], ['c']]`.

    В Excel для установки матричной привязки может использоваться любое связанное выделение ячеек. В Word матричная привязка поддерживается только таблицами.

1. **[Табличная привязка][TableBinding]**. Выполняет привязку к области документа, содержащей таблицу с заголовками. Данные в табличной привязке записываются или считываются как объект [TableData](/javascript/api/office/office.tabledata). Объект `TableData` предоставляет данные с помощью свойств `headers` и `rows`.

    Любая таблица Excel или Word может быть основой для табличной привязки. После создания табличной привязки каждая новая строка или столбец, добавляемые пользователем в таблицу, автоматически включаются в привязку.

После создания привязки с помощью одного из трех методов addFrom `Bindings` объекта можно работать с данными и свойствами привязки с помощью методов соответствующего объекта: [MatrixBinding], [TableBinding] или [TextBinding]. Все три объекта наследуют методы [getDataAsync] и [setDataAsync] объекта `Binding`, позволяющие работать со связанными данными.

> [!NOTE]
> **В каких случаях следует использовать матричные и табличные привязки?** Если табличные данные, с которыми вы работаете, содержат строку итогов, а сценарию надстройки необходимо получить доступ к значениям в этой строке или проверить, находится ли в ней выбранный пользователем фрагмент, то необходимо использовать матричную привязку. Если установить привязку к табличным данным, содержащим строку итогов, то значения свойства [TableBinding.rowCount], а также свойств `rowCount` и `startRow` объекта [BindingSelectionChangedEventArgs] в обработчиках событий не будут отражать строку итогов. Чтобы обойти это ограничение, необходимо установить матричную привязку для работы со строкой итогов.

## <a name="add-a-binding-to-the-users-current-selection"></a>Добавление привязки к текущему фрагменту, выделенному пользователем

В приведенном ниже примере показано, как добавить текстовую привязку с именем `myBinding` к текущему выделенному фрагменту в документе с помощью метода [addFromSelectionAsync].

```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В этом примере указанным типом привязки является текст. Это означает, что для выделенного фрагмента будет создан объект [TextBinding]. Различные типы привязок предоставляют различные данные и операции. [Office.BindingType] — это перечисление доступных значений типов привязки.

Вторым дополнительным параметром является объект, который указывает идентификатор новой создаваемой привязки. Если идентификатор не указан, он создается автоматически.

Анонимная функция, которая передается в метод в качестве  последнего параметра обратного вызова, выполняется по завершении создания привязки. Функция вызывается с использованием параметра `asyncResult`, предоставляющего доступ к объекту [AsyncResult], который сообщает состояние вызова. Свойство `AsyncResult.value` содержит ссылку на объект [Binding] того типа, который указан для новой привязки. С помощью объекта [Binding] можно получать и задавать данные.

## <a name="add-a-binding-from-a-prompt"></a>Добавление привязки по запросу

В приведенном ниже примере показано, как добавить текстовую привязку с именем `myBinding`, используя метод [addFromPromptAsync]. Этот метод позволяет пользователю указать диапазон для привязки с помощью встроенного в приложение запроса на выбор диапазона.

```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

В этом примере указанным типом привязки является текст. Это означает, что для выделенного фрагмента, указанного пользователем в окне ввода, будет создан объект [TextBinding].

Вторым параметром является объект, который содержит идентификатор новой создаваемой привязки. Если идентификатор не указан, он создается автоматически.

Анонимная функция, переданная в метод в  качестве третьего параметра обратного вызова, выполняется после создания привязки. При выполнении функции обратного вызова объект [AsyncResult] содержит сведения о состоянии вызова и только что созданную привязку.

На рис. 1 показано встроенное окно запроса выбора диапазона в Excel.

*Рис. 1. Пользовательский интерфейс выбора данных в Excel*

![Диалоговое окно "Выбор данных".](../images/agave-api-overview-excel-selection-ui.png)

## <a name="add-a-binding-to-a-named-item"></a>Добавление привязки к именованному элементу

`myRange` В следующем примере показано, как добавить привязку к существующему именованному элементу в качестве "матрицы" привязки с помощью метода [addFromNamedItemAsync] и `id` присвоить привязке значение myMatrix.

```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

**Для Excel**`itemName` параметр метода [addFromNamedItemAsync] может ссылаться на существующий именованный диапазон, `A1` `("A1:A3")`диапазон, указанный со стилем ссылки, или таблицу. По умолчанию при добавлении таблиц в Excel имя "Table1" назначается первой добавленной таблице, "Table2" — второй таблице и так далее. Чтобы назначить понятное имя для таблицы в пользовательском интерфейсе Excel, `Table Name` используйте свойство в таблице **| Вкладка** конструктора ленты.

> [!NOTE]
> В Excel при указании таблицы в качестве именованного элемента необходимо полностью указать имя, чтобы включить имя листа в имя таблицы в следующем формате: `"Sheet1!Table1"`

В следующем примере создается привязка в Excel к первым трем ячейкам в столбце A ( `"A1:A3"`), `"MyCities"`присваиваются идентификаторы, а затем в эту привязку записывают три названия городов.

```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

**Для Word** параметр `itemName` метода [addFromNamedItemAsync] `Title` ссылается на свойство элемента управления `Rich Text` содержимым. (`Rich Text` — единственный элемент управления содержимым, поддерживающий привязку.)

По умолчанию элементу управления содержимым не назначено `Title*`значение. Чтобы назначить понятное имя в пользовательском интерфейсе Word, после вставки элемента управления контентом **Форматированный текст** из группы **Элементы управления** на вкладке **Разработчик** ленты выберите команду **Свойства** в группе **Элементы управления**, чтобы открыть диалоговое окно **Свойства элемента управления контентом**. Затем задайте `Title` для свойства элемента управления содержимым имя, на которое вы хотите ссылаться из кода.

В следующем примере создается текстовая привязка в Word `"FirstName"`к элементу управления содержимым форматированного текста с именем, назначается **идентификатор**`"firstName"`, а затем отображаются эти сведения.

```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="get-all-bindings"></a>Получение всех привязок

В приведенном ниже примере показано, как получить все привязки в документе с помощью метода Bindings.[getAllAsync].

```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    let bindingString = '';
    for (let i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Анонимная функция, передаваемая в метод в `callback` качестве параметра, выполняется после завершения операции. Функция вызывается с одним параметром, `asyncResult`который содержит массив привязок в документе. Массив перебирается для создания строки, содержащей идентификаторы привязок. Строка отображается в окне сообщения.

## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a>Получение привязки по идентификатору с помощью метода getByIdAsync объекта Bindings

В приведенном ниже примере показано, как с помощью метода [getByIdAsync] получить привязку в документе, указав ее идентификатор. В этом примере предполагается, что привязка с именем `'myBinding'` была добавлена в документ с помощью одного из методов, описанных ранее в этой статье.

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В этом примере первый `id` параметр — это идентификатор извлекаемой привязки.

Анонимная функция, которая передается в метод в качестве  второго параметра обратного вызова, выполняется после завершения операции. Функция вызывается с передачей одного параметра _asyncResult_, который содержит состояние вызова и привязки с идентификатором "myBinding".

## <a name="get-a-binding-by-id-using-the-select-function-of-the-office-object"></a>Получение привязки по идентификатору с помощью функции select объекта Office

В следующем примере показано, как с помощью функции [Office.select] получить обещание объекта [привязки] в документе, указав его идентификатор в строке выбора. Затем вызывается метод Binding.[getDataAsync] для получения данных из указанной привязки. В этом примере предполагается, что привязка с именем `'myBinding'` была добавлена в документ с помощью одного из методов, описанных ранее в этой статье.

```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

> [!NOTE]
> `select` Если обещание функции успешно возвращает объект [Binding], этот объект предоставляет только следующие четыре метода объекта: [getDataAsync], [setDataAsync], [addHandlerAsync] и [removeHandlerAsync]. Если обещание не может вернуть объект Binding, `onError` обратный вызов можно использовать для доступа к объекту [asyncResult.error] для получения дополнительных сведений. Если необходимо вызвать член объекта Binding, отличный от четырех методов, предоставляемых обещанием объекта [Binding]`select`, возвращенным функцией, вместо этого используйте метод [getByIdAsync] с помощью свойства [Document.bindings] и привязок.[ Метод getByIdAsync] для получения объекта [Binding].

## <a name="release-a-binding-by-id"></a>Отмена привязки по идентификатору

В приведенном ниже примере показано, как с помощью метода [releaseByIdAsync] удалить привязку из документа, указав ее идентификатор.

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В этом примере первый параметр `id` — идентификатор удаляемой привязки.

Анонимная функция, которая передается в метод в качестве второго параметра, является обратным вызовом, который выполняется после завершения операции. Функция вызывается с передачей одного параметра [asyncResult], который содержит состояние вызова.

## <a name="read-data-from-a-binding"></a>Чтение данных из привязки

В приведенном ниже примере показано, как с помощью метода [getDataAsync] получить данные из существующей привязки.

```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

`myBinding` — переменная, содержащая существующую текстовую привязку в документе. Кроме того, с помощью метода [Office.select] можно получить доступ к привязке по ее идентификатору и вызвать метод [getDataAsync]. Вот как это сделать:

```js
Office.select("bindings#myBindingID").getDataAsync
```

Анонимная функция, передаваемая в метод, является обратным вызовом, который выполняется после завершения операции. Свойство [AsyncResult].value содержит данные в `myBinding`. Тип значения зависит от типа привязки. В этом примере используется привязка текста, поэтому значение должно содержать строку. Следовательно, значение будет содержать строку. Дополнительные примеры работы с матричными и табличными привязками представлены в статье, посвященной методу [getDataAsync].

## <a name="write-data-to-a-binding"></a>Запись данных в привязку

В приведенном ниже примере показано, как с помощью метода [setDataAsync] задать данные в существующей привязке.

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

`myBinding` — переменная, содержащая существующую текстовую привязку в документе.

В этом примере первый параметр — это значение, которое необходимо задать `myBinding`. Так как привязка текстовая, этим значением будет `string`. Привязки разных типов принимают разные типы данных.

Анонимная функция, передаваемая в метод, является обратным вызовом, который выполняется после завершения операции. Функция вызывается с одним параметром, `asyncResult`который содержит состояние результата.

> [!NOTE]
> С момента выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel в Интернете можно [задавать форматирование при записи или обновлении данных в связанных таблицах](../excel/excel-add-ins-tables.md).

## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a>Обнаружение изменений в данных или выделенном фрагменте для привязки

В приведенном ниже примере показано, как присоединить обработчик событий к событию [DataChanged](/javascript/api/office/office.binding) привязки с идентификатором MyBinding.

```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

`myBinding` — переменная, содержащая существующую текстовую привязку в документе.

Первый параметр _eventType_ метода [addHandlerAsync] указывает имя события для подписки. [Office.EventType] — это перечисление доступных значений типов событий. `Office.EventType.BindingDataChanged` вычисляет строку bindingDataChanged.

Функция`dataChanged`, которая передается в метод в качестве второго параметра  обработчика, является обработчиком событий, который выполняется при изменении данных в привязке. Функция вызывается с одним параметром _eventArgs_, который содержит ссылку на привязку. Эта привязка может использоваться для получения обновленных данных.

Так же вы можете определять, поменял ли пользователь выделенный фрагмент в привязке, добавив обработчик события [SelectionChanged] привязки. Для этого задайте параметр `eventType` метода [addHandlerAsync] как `Office.EventType.BindingSelectionChanged` или `"bindingSelectionChanged"`.

Вы можете добавить несколько обработчиков событий для этого события, снова вызвав метод [addHandlerAsync] и передав дополнительную функцию обработчика событий для параметра `handler`. Это возможно при условии, что имя каждой функции обработчика событий уникально.

### <a name="remove-an-event-handler"></a>Удаление обработчика события

Чтобы удалить обработчик какого-либо события, вызовите метод [removeHandlerAsync], передав тип события в качестве первого параметра _eventType_, а имя удаляемой функции обработчика событий — в качестве второго параметра _handler_. Например, приведенная ниже функция удалит функцию обработчика событий `dataChanged`, добавленную в примере, который представлен в предыдущем разделе.

```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```

> [!IMPORTANT]
> Если _необязательный_ параметр обработчика опущен при вызове метода [removeHandlerAsync] , `eventType` все обработчики событий для указанного объекта будут удалены.

## <a name="see-also"></a>Дополнительные ресурсы

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [Асинхронное программирование в надстройках для Office](asynchronous-programming-in-office-add-ins.md)
- [Выполняйте чтение и запись данных при активном выделении фрагмента в документе или электронной таблице.](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)

[Binding]:               /javascript/api/office/office.binding
[MatrixBinding]:         /javascript/api/office/office.matrixbinding
[TableBinding]:          /javascript/api/office/office.tablebinding
[TextBinding]:           /javascript/api/office/office.textbinding
[getDataAsync]:          /javascript/api/office/office.binding#getDataAsync_options__callback_
[setDataAsync]:          /javascript/api/office/office.binding#setDataAsync_data__options__callback_
[SelectionChanged]:      /javascript/api/office/office.bindingselectionchangedeventargs
[addHandlerAsync]:       /javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_
[removeHandlerAsync]:    /javascript/api/office/office.binding#removeHandlerAsync_eventType__options__callback_

[Bindings]:              /javascript/api/office/office.bindings
[getByIdAsync]:          /javascript/api/office/office.bindings#getByIdAsync_id__options__callback_
[getAllAsync]:           /javascript/api/office/office.bindings#getAllAsync_options__callback_
[addFromNamedItemAsync]: /javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_
[addFromSelectionAsync]: /javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_
[addFromPromptAsync]:    /javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_
[releaseByIdAsync]:      /javascript/api/office/office.bindings#releaseByIdAsync_id__options__callback_

[AsyncResult]:          /javascript/api/office/office.asyncresult
[Office.BindingType]:   /javascript/api/office/office.bindingtype
[Office.select]:        /javascript/api/office 
[Office.EventType]:     /javascript/api/office/office.eventtype 
[Document.bindings]:    /javascript/api/office/office.document

[TableBinding.rowCount]: /javascript/api/office/office.tablebinding
[BindingSelectionChangedEventArgs]: /javascript/api/office/office.bindingselectionchangedeventargs