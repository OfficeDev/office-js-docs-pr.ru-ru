
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a>Привязка к областям в документе или электронной таблице

Благодаря доступу к данным на основе привязок контентные надстройки и надстройки области задач могут согласованно получать доступ к определенной области документа или электронной таблицы с помощью идентификатора. Прежде всего надстройке необходимо создать привязку, вызвав один из методов, сопоставляющих часть документа с уникальным идентификатором: [addFromPromptAsync], [addFromSelectionAsync] или [addFromNamedItemAsync]. После создания привязки надстройка может использовать предоставленный идентификатор для доступа к данным, содержащимся в сопоставленной области документа или электронной таблицы. При создании привязки надстройке предоставляется указанное ниже значение.


- Разрешает доступ к общим структурам данных в поддерживаемых приложениях Office, таким как: таблицы, диапазоны или текст (связанная последовательность знаков).
    
- Позволяет производить операции чтения или записи без необходимости выделения пользователем фрагмента.
    
- Устанавливает отношение между надстройкой и данными в документе. Привязки сохраняются в документе и могут использоваться позже.
    
Установка привязки также позволяет подписываться на данные и выбирать изменения событий, относящиеся к конкретной области документа или электронной таблицы. Это означает, что надстройка уведомляется только об изменениях, происходящих внутри данной конкретной области, в отличие от изменений, затрагивающих в целом весь документ или электронную таблицу.

Объект [Bindings] предоставляет метод [getAllAsync], открывающий доступ к полному набору привязок, установленных в документе или на листе. Доступ к отдельной привязке можно получить по ее идентификатору с помощью метода Bindings.[getByIdAsync] или [Office.select]. Вы можете устанавливать новые привязки или удалять существующие с помощью одного из следующих методов объекта [Bindings]: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync] или [releaseByIdAsync].


## <a name="binding-types"></a>Типы привязок

[Привязки ][Office.BindingType] бывают трех типов. Такой тип вы можете задать с помощью параметра _bindingType_, когда создаете привязку с использованием метода [addFromSelectionAsync], [addFromPromptAsync] или [addFromNamedItemAsync]

1. **[Текстовая привязка][TextBinding]**. Выполняет привязку к области документа, которая может быть представлена как текст.

    В Word поддерживается большинство связанных выделений, тогда как в Excel для привязки текста можно использовать только выделения отдельных ячеек. Excel поддерживает только обычный текст, а Word — три формата: обычный текст, HTML и Open XML для Office.

2. **[Матричная привязка][MatrixBinding]**. Выполняет привязку к фиксированной области документа, содержащей табличные данные без заголовков. Данные в матричной привязке записываются или считываются как двумерный объект **Array**, который реализуется в JavaScript как массив массивов. Например, две строки значений типа **string** в двух столбцах можно записывать или считывать как ` [['a', 'b'], ['c', 'd']]`, а один столбец из трех строк можно записывать или считывать как `[['a'], ['b'], ['c']]`.

    В Excel для установки матричной привязки можно использовать любое связанное выделение ячеек. В Word матричную привязку поддерживают только таблицы.

3. **[Табличная привязка][TableBinding]**. Выполняет привязку к области документа, содержащей таблицу с заголовками. Данные в табличной привязке записываются или считываются как объект [TableData](http://dev.office.com/reference/add-ins/shared/tabledata). Объект `TableData` предоставляет данные с помощью свойств `headers` и `rows`.

    Любая таблица Excel или Word может быть основой для табличной привязки. После создания табличной привязки каждая новая строка или столбец, добавляемые пользователем в таблицу, автоматически включаются в привязку.

Создав привязку с помощью одного из трех методов addFrom объекта `Bindings`, вы можете работать с данными и свойствами привязки с помощью методов соответствующего объекта: [MatrixBinding], [TableBinding] или [TextBinding]. Все три объекта наследуют методы [getDataAsync] и [setDataAsync] объекта `Binding`, позволяющие работать со связанными данными.

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

Анонимная функция, передаваемая функции в качестве конечного параметра _callback_, выполняется по завершении создания привязки. Функция вызывается с использованием параметра `asyncResult`, предоставляющего доступ к объекту [AsyncResult], который сообщает состояние вызова. Свойство `AsyncResult.value` содержит ссылку на объект [Binding] того типа, который указан для новой привязки. С помощью объекта [Binding] можно получать и задавать данные.

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

Анонимная функция, которая передается в функцию в качестве третьего параметра _callback_, выполняется после создания привязки. При выполнении функции обратного вызова объект [AsyncResult] содержит сведения о состоянии вызова и только что созданную привязку.

На рис. 1 показано встроенное окно запроса выбора диапазона в Excel.


**Рис. 1. Пользовательский интерфейс выбора данных в Excel**

![Пользовательский интерфейс выбора данных в Excel](../images/AgaveAPIOverview_ExcelSelectionUI.png)


## <a name="add-a-binding-to-a-named-item"></a>Добавление привязки к именованному элементу


В приведенном ниже примере показано, как добавить матричную привязку в существующий именованный элемент `myRange`, используя метод [addFromNamedItemAsync], и назначить параметру `id` этой привязки значение myMatrix.


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

 **В случае Excel** параметр `itemName` метода [addFromNamedItemAsync] может ссылаться на диапазон, указанный с использованием ссылки вида `A1` `("A1:A3")`, существующий именованный диапазон или таблицу. По умолчанию при добавлении таблиц в Excel первой таблице назначается имя "Таблица1", второй — "Таблица2" и т. д. Назначить таблице понятное имя в пользовательском интерфейсе Excel можно с помощью свойства **Имя таблицы** на вкладке **Работа с таблицами | Конструктор** на ленте.


 >**Примечание.** В Excel при задании таблицы в качестве именованного элемента необходимо указать ее имя полностью, включая имя листа, в таком формате: `"Sheet1!Table1"`.

В приведенном ниже примере показано, как в Excel создать привязку к первым трем ячейкам столбца A (`"A1:A3"`), назначить значение id `"MyCities"`, а затем записать три названия города в эту привязку.


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

 **В случае Word** параметр `itemName` метода [addFromNamedItemAsync] ссылается на свойство `Title`, принадлежащее элементу управления содержимым `Rich Text`. (`Rich Text` — единственный элемент управления содержимым, поддерживающий привязку.)

По умолчанию элементу управления содержимым не назначено значение `Title*`. Чтобы назначить понятное имя в пользовательском интерфейсе Word, выполните следующее. Вставьте элемент управления содержимым **Форматированный текст** группу **Элементы управления** на вкладке **Разработчик** ленты. Выберите команду **Свойства** в группе **Элементы управления**, чтобы открыть диалоговое окно **Свойства элемента управления содержимым**. Задайте для свойства **Title**, принадлежащего элементу управления содержимым, имя, на которое вы будете ссылаться в коде.

В следующем примере показано, как в Word создать привязку текста к элементу управления контентом "Форматированный текст" с именем `"FirstName"`, назначить  **id**`"firstName"`, а затем отобразить эти сведения.


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
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Анонимная функция, передаваемая функции в качестве параметра `callback`, выполняется по завершении операции. Функция вызывается с использованием параметра `asyncResult`, содержащего массив привязок в документе. Выполняется итерация массива для составления строки, содержащей идентификаторы привязок. Затем эта строка отображается в окне сообщения.


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

В этом примере первый параметр `id` — идентификатор получаемой привязки.

Анонимная функция, которая передается в функцию в качестве второго параметра _callback_. Функция вызывается с передачей одного параметра _asyncResult_, который содержит состояние вызова и привязки с идентификатором "myBinding".


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a>Получение привязки по идентификатору с помощью метода select объекта Office


В приведенном ниже примере показано, как с помощью метода [Office.select] получить обещание для объекта [Binding] в документе, указав его идентификатор в строке селектора. Затем вызывается метод Binding.[getDataAsync] для получения данных из указанной привязки. В этом примере предполагается, что привязка с именем `'myBinding'` была добавлена в документ с помощью одного из методов, описанных ранее в этой статье.


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


 > **Примечание.**  Если при получении обещания для метода `select` успешно возвращается объект [Binding], то этот объект предоставляет только следующие четыре метода: [getDataAsync], [setDataAsync], [addHandlerAsync] и [removeHandlerAsync]. Если же возвратить объект Binding не удается, то для получения дополнительной информации можно получить доступ к объекту [asyncResult].error с помощью параметра обратного вызова `onError`. Если вам необходимо вызвать элемент объекта Binding, которого нет среди четырех методов, предоставленных с обещанием объекта Binding, которое возвращено методом `select`, примените метод [getByIdAsync]. Для этого с помощью свойства [Document.bindings] и метода Bindings.[getByIdAsync] получите объект Binding**.

## <a name="release-a-binding-by-id"></a>Удаление привязки по идентификатору


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

Анонимная функция, которая передается в функцию в качестве второго параметра, является функцией обратного вызова, которая выполняется после завершения операции. Функция вызывается с передачей одного параметра [asyncResult], который содержит состояние вызова.


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


Анонимная функция передается функции в качестве параметра callback и выполняется по завершении операции. Свойство [AsyncResult].value содержит данные в `myBinding`. Тип значения зависит от типа привязки. В этом примере используется текстовая привязка. Следовательно, значение будет содержать строку. Дополнительные примеры работы с матричными и табличными привязками представлены в статье, посвященной методу [getDataAsync].


## <a name="write-data-to-a-binding"></a>Запись данных в привязку

В приведенном ниже примере показано, как с помощью метода [setDataAsync] задать данные в существующей привязке.

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 `myBinding` — переменная, содержащая существующую текстовую привязку в документе.

В этом примере первый параметр — значение, задаваемое для `myBinding`. Так как привязка текстовая, этим значением будет `string`. Привязки разных типов принимают разные типы данных.

Анонимная функция передается функции в качестве параметра callback и выполняется по завершении операции. Функция вызывается с использованием параметра `asyncResult`, содержащего состояние результата.

 > **Примечание.** С момента выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel Online можно [задавать форматирование при записи или обновлении данных в связанных таблицах](../../docs/excel/format-tables-in-add-ins-for-excel.md).


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a>Обнаружение изменений в данных или выделении в привязке


В примере ниже показано, как присоединить обработчик события к событию [DataChanged](http://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) привязки с идентификатором MyBinding.


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

Первый параметр `eventType` метода [addHandlerAsync] задает имя события для подписки. [Office.EventType] — это перечисление доступных значений типов событий. `Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"`.

Функция `dataChanged`, которая передается в эту функцию в качестве второго параметра _handler_, является вторым обработчиком событий, который выполняется при изменении данных в привязке. Функция вызывается с одним параметром _eventArgs_, который содержит ссылку на привязку. Эта привязка может использоваться для получения обновленных данных.

Так же вы можете определять, поменял ли пользователь выделенный фрагмент в привязке, добавив обработчик события [SelectionChanged] привязки. Для этого задайте параметр `eventType` метода [addHandlerAsync] как `Office.EventType.BindingSelectionChanged` или `"bindingSelectionChanged"`.

Вы можете добавить несколько обработчиков событий для этого события, снова вызвав метод [addHandlerAsync] и передав дополнительную функцию обработчика событий для параметра `handler`. Это возможно при условии, что имя каждой функции обработчика событий уникально.


### <a name="remove-an-event-handler"></a>Удаление обработчика события


Чтобы удалить обработчик какого-либо события, вызовите метод [removeHandlerAsync], передав тип события в качестве первого параметра _eventType_, а имя удаляемой функции обработчика событий — в качестве второго параметра _handler_. Например, приведенная ниже функция удалит функцию обработчика событий `dataChanged`, добавленную в примере, который представлен в предыдущем разделе.


```
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


 >**Важно!**  Если при вызове метода [removeHandlerAsync] не указать необязательный параметр _handler_, то все обработчики событий для указанного параметра `eventType` будут удалены.


## <a name="additional-resources"></a>Дополнительные ресурсы

- [Общие сведения об интерфейсе API JavaScript для Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Асинхронное программирование в надстройках для Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Выполняйте чтение и запись данных при активном выделении фрагмента в документе или электронной таблице.](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
[Binding]:               ../../reference/shared/binding.md
[MatrixBinding]:         ../../reference/shared/binding.matrixbinding.md
[TableBinding]:          ../../reference/shared/binding.tablebinding.md
[TextBinding]:           ../../reference/shared/binding.textbinding.md
[getDataAsync]:          ../../reference/shared/binding.getdataasync.md
[setDataAsync]:          ../../reference/shared/binding.setdataasync.md
[SelectionChanged]:      ../../reference/shared/binding.bindingselectionchangedevent.md
[addHandlerAsync]:       ../../reference/shared/binding.addhandlerasync.md
[removeHandlerAsync]:    ../../reference/shared/binding.removehandlerasync.md

[Bindings]:              ../../reference/shared/bindings.bindings.md
[getByIdAsync]:          ../../reference/shared/bindings.getbyidasync.md 
[getAllAsync]:           ../../reference/shared/bindings.getallasync.md
[addFromNamedItemAsync]: ../../reference/shared/bindings.addfromnameditemasync.md
[addFromSelectionAsync]: ../../reference/shared/bindings.addfromselectionasync.md
[addFromPromptAsync]:    ../../reference/shared/bindings.addfrompromptasync.md
[releaseByIdAsync]:      ../../reference/shared/bindings.releasebyidasync.md

[AsyncResult]:          ../../reference/shared/asyncresult.md
[Office.BindingType]:   ../../reference/shared/bindingtype-enumeration.md
[Office.select]:        ../../reference/shared/office.select.md 
[Office.EventType]:     ../../reference/shared/eventtype-enumeration.md 
[Document.bindings]:    ../../reference/shared/document.bindings.md


[TableBinding.rowCount]: ../../reference/shared/binding.tablebinding.rowcount.1md
[BindingSelectionChangedEventArgs]: ../../reference/shared/binding.bindingselectionchangedeventargs.md
