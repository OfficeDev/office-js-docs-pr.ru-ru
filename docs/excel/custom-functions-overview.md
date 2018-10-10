---
ms.date: 09/27/2018
description: Создание настраиваемой функции в Excel с помощью JavaScript.
title: Создание настраиваемых функций в Excel (предварительная версия)
ms.openlocfilehash: f6b658bbd119a785b342ec22bc1b341f6902da3f
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459345"
---
# <a name="create-custom-functions-in-excel-preview"></a>Создание настраиваемых функций в Excel (предварительная версия)

Настраиваемые функции позволяют разработчикам добавлять новые функции в Excel, определяя эти функции в JavaScript как часть надстройки. Пользователи в Excel могут получать доступ к настраиваемым функциям так же, как к любой собственной функции в Excel, например `SUM()`. В этой статье описывается, как создавать настраиваемые функции в Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Следующий рисунок демонстрирует процесс вставки настраиваемой функции в рабочий лист Excel конечным пользователем. Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которые пользователь указывает в качестве входных параметров для функции.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

Следующий код определяет настраиваемую функцию `ADD42`.

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> В разделе [Известные проблемы](#known-issues) далее в этой статье указаны текущие ограничения настраиваемых функций.

## <a name="components-of-a-custom-functions-add-in-project"></a>Компоненты проекта надстройки пользовательских функций

Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания проекта надстройки настраиваемых функций Excel, вы увидите следующие файлы в проекте, который создает генератор:

| Файл | Формат файла | Описание |
|------|-------------|-------------|
| **./src/customfunctions.js**<br/>или<br/>**./src/customfunctions.ts** | JavaScript<br/>или<br/>TypeScript | Содержит код, который определяет настраиваемые функции. |
| **./config/customfunctions.json** | JSON | Содержит метаданные, которые описывают настраиваемые функции и позволяют Excel регистрировать настраиваемые функции, чтобы сделать их доступными для пользователей. |
| **./index.html** | HTML | Предоставляет ссылку в тегах &lt;script&gt; на файл JavaScript, который определяет пользовательские функции. |
| **./manifest.xml** | XML | Указывает пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML, указанных ранее в этой таблице. |

Дополнительные сведения об этих файлах можно найти в следующих разделах.

### <a name="script-file"></a>Файл сценария 

Файл сценария (**./src/customfunctions.js** или **./src/customfunctions.ts** в проекте, который создает генератор Yo Office) содержит код, который определяет настраиваемые функции и сопоставляется с объектами в [файле метаданных JSON](#json-metadata-file). 

Например, следующий код определяет настраиваемые функции `add` и `increment`, а затем определяет информацию о сопоставлении для обеих функций. Функция `add` сопоставляется с объектом в файле метаданных JSON, где значение свойства `id` равно **ADD**, а функция `increment` сопоставляется с объектом в файле метаданных, где значение свойства `id` равно **INCREMENT**. Подробнее о сопоставлении имен функций в файле сценария с объектами в файле метаданных JSON см. [Практические рекомендации по настраиваемым функциям](custom-functions-best-practices.md#mapping-function-names-to-json-metadata).

```js
function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

// map `id` values in the JSON metadata file to the JavaScript function names
CustomFunctionMappings.ADD = add;
CustomFunctionMappings.INCREMENT = increment;
```

### <a name="json-metadata-file"></a>Файл метаданных JSON 

Файл метаданных настраиваемых функций (**./config/customfunctions.json** в проекте, создаваемом генератором Yo Office) предоставляет информацию, которую Excel требует, чтобы зарегистрировать настраиваемые функции и сделать их доступными для конечных пользователей. Настраиваемые функции регистрируются, когда пользователь запускает надстройку в первый раз. После этого они доступны для того же пользователя во всех книгах (т. е. не только в книге, в которой первоначально выполнялась надстройка).

> [!TIP]
> Чтобы настраиваемые функции правильно работали в Excel Online, в параметры сервера, на котором размещен файл JSON, необходимо включить [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS).

Следующий код в **customfunctions.json** определяет метаданные для описанных ранее функций `add` и `increment`. В таблице, следующей за данным примером кода, приведены подробные сведения об отдельных свойствах в этом объекте JSON. См. [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) для получения более подробных сведений о задании значений для свойств `id` и `name` в файле метаданных JSON.

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      "description": "Periodically increment a value",
      "helpUrl": "http://www.contoso.com",
      "result": {
          "type": "number",
          "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "increment",
            "description": "Amount to increment",
            "type": "number",
            "dimensionality": "scalar"
        }
    ],
    "options": {
        "cancelable": true,
        "stream": true
      }
    }
  ]
}
```

В следующей таблице перечислены свойства, которые обычно присутствуют в файле метаданных JSON. Более подробные сведения о файле метаданных JSON см. в статье [Метаданные настраиваемых функций](custom-functions-json.md).

| Свойство  | Описание |
|---------|---------|
| `id` | Уникальный идентификатор для функции. Изменение этого идентификатора после его установки не допускается. |
| `name` | Имя функции, которое конечный пользователь видит в Excel. В Excel название этой функции будет иметь префикс пространства имен настраиваемых функций, который указан в [XML-файле манифеста](#manifest-file). |
| `helpUrl` | URL-адрес страницы, которая отображается, когда пользователь запрашивает справку. |
| `description` | Описывает, что выполняет функция. Это значение появляется как подсказка, когда функция является выбранным элементом в меню автозаполнения в Excel. |
| `result`  | Объект, который определяет тип данных, возвращаемых функцией. Значение дочернего свойства `type` может быть **string**, **number** или **boolean**. Дочернему свойству `dimensionality` может присваиваться значение **scalar** или **matrix** (двумерный массив значений указанного типа `type`). |
| `parameters` | Массив, который определяет входные параметры для функции. Дочерние свойства `name` и `description` отображаются в Excel intelliSense. Значение дочернего свойства `type` может быть **string**, **number** или **boolean**. Дочернему свойству `dimensionality` может присваиваться значение **scalar** или **matrix** (двумерный массив значений указанного типа `type`). |
| `options` | Позволяет настроить некоторые аспекты того, как и когда Excel выполняет эту функцию. Подробнее о том, как это свойство можно использовать, см. в разделах [Потоковые функции](#streaming-functions) и [Отмена функции](#canceling-a-function) ниже в этой статье. |

### <a name="manifest-file"></a>Файл манифеста

XML-файл манифеста для надстройки, который определяет настраиваемые функции (**./manifest.xml** в проекте, создаваемом генератором Yo Office), определяет пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML. Ниже показан пример использования элементов `<ExtensionPoint>` и `<Resources>` в разметке XML. Эти элементы необходимо включить в манифест надстройки, чтобы иметь возможность выполнять настраиваемые функции.  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Функции Excel добавляются пространством имен, указанным в файле манифеста XML. Пространство имен функции предшествует имени функции и отделяется от него точкой. Например, чтобы вызвать функцию `ADD42` в ячейке листа Excel, следует ввести `=CONTOSO.ADD42`, так как CONTOSO — это пространство имен и `ADD42` — это имя функции, указанной в файле JSON. Пространство имен предназначено для использования в качестве идентификатора для вашей компании или надстройки. 

## <a name="functions-that-return-data-from-external-sources"></a>Функции, возвращающие данные из внешних источников

Если настраиваемая функция получает данные из внешнего источника, например веб-сайта, она должна:

1. возвращать обещание JavaScript в Excel.

2. разрешать Promise окончательным значением, используя функцию обратного вызова.

Пока Excel ожидает конечный результат, настраиваемые функции отображают в ячейке временный результат `#GETTING_DATA`. Во время ожидания результата пользователи могут нормально взаимодействовать с остальной частью листа.

В следующем примере кода настраиваемая функция `getTemperature()` получает от термометра текущую температуру. Обратите внимание на то, что функция `sendWebRequest` является гипотетической (не указывается здесь) и использует [XHR](custom-functions-runtime.md#xhr-example) для вызова веб-службы температуры.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a>Потоковые функции

Потоковые настраиваемые функции позволяют вам выводить данные в ячейки многократно с течением времени, не требуя от пользователя явно запрашивать обновление данных. Следующий пример кода представляет собой настраиваемую функцию, которая каждую секунду добавляет число к результату. Обратите внимание на следующие особенности этого кода:

- Excel автоматически отображает каждое новое значение при помощи обратного вызова `setResult`.

- Второй входной параметр `handler` не отображается для конечных пользователей в Excel при выборе функции из меню автозаполнения.

- Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции. Для любой потоковой функции необходимо реализовать подобный обработчик отмены. Подробнее см. раздел [Отмена функции](#canceling-a-function). 

```js
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}
```

При указании метаданных для потоковой функции в файле метаданных JSON необходимо задать свойства `"cancelable": true` и `"stream": true` для объекта `options`, как показано в следующем примере.

```json
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a>Отмена функции

В некоторых случаях может потребоваться отменить выполнение потоковой настраиваемой функции, чтобы снизить потребление ею пропускной способности, рабочей памяти и загрузку процессора. Excel отменяет выполнение функции в следующих ситуациях.

- Когда пользователь редактирует или удаляет ячейку, содержащую ссылку на функцию.

- Когда изменяется один из аргументов (входных параметров) функции. В этом случае после отмены активируется новый вызов функции.

- Когда пользователь запускает пересчет вручную. В этом случае после отмены активируется новый вызов функции.

Чтобы включить возможность отмены функции, необходимо реализовать обработчик отмены в функции JavaScript и указать свойство `"cancelable": true` в объекте `options` в метаданных JSON, которые описывают функцию. В примерах кода в предыдущем разделе данной статьи приводится пример этой техники.

## <a name="saving-and-sharing-state"></a>Сохранение и передача состояния

Настраиваемые функции могут сохранять данные в глобальных переменных JavaScript. При последующих вызовах настраиваемая функция может использовать значения, сохраненные в этих переменных. Сохранение состояния может быть полезно, когда пользователи добавляют одну настраиваемую функцию к нескольким ячейкам, потому что все экземпляры функции могут совместно использовать ее состояние. Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось делать дополнительные вызовы одного и того же веб-ресурса.

В приведенном ниже примере кода показана реализация вышеописанной потоковой функции температуры, осуществляющей глобальное сохранение состояния. Обратите внимание на следующие особенности этого кода:

- `refreshTemperature`  — это потоковая функция, ежесекундно считывающая температуру определенного термометра. Новые температуры сохраняются в переменную `savedTemperatures`, но не обновляют значение ячейки напрямую. Она не должна вызываться непосредственно из ячейки листа, * поэтому она не регистрируется в файле JSON*.

- `streamTemperature` обновляет значения температуры, которые отображаются в ячейке каждую секунду, а в качестве источника данных использует переменную `savedTemperatures`. Она должна быть зарегистрирована в файле JSON и записана прописными буквами: `STREAMTEMPERATURE`.

- Пользователи могут вызывать функцию `streamTemperature` из нескольких ячеек в пользовательском Интерфейсе Excel. Каждый вызов считывает данные из той же переменной `savedTemperatures`.

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a>Работа с диапазонами данных

Настраиваемая функция может принимать диапазон данных в качестве входного параметра, или она может возвращать диапазон данных. В JavaScript диапазон данных представляется как двухмерный массив.

Предположим, к примеру, что ваша функция возвращает второе наибольшее значение из диапазона чисел, хранящихся в Excel. Следующая функция принимает параметр `values`, который имеет тип `Excel.CustomFunctionDimensionality.matrix`. Обратите внимание, что в метаданных JSON для этой функции вы должны для параметра `type` установить значение `matrix`.

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="handling-errors"></a>Обработка ошибок

При построении надстройки, определяющей настраиваемые функции, не забудьте добавить логику для обработки ошибок, возникающих в среде выполнения. Обработка ошибок для настраиваемых функций такая же, как и в случае [обработки ошибок для API JavaScript Excel в целом](excel-add-ins-error-handling.md). В следующем примере кода метод `.catch` будет обрабатывать все ошибки, возникающие ранее в коде.

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="known-issues"></a>Известные проблемы

- URL-адреса справки и описания параметров пока не используются в Excel.
- Настраиваемые функции в настоящее время недоступны в Excel для мобильных клиентов.
- Изменяемые функции (которые пересчитываются автоматически при изменении несвязанных данных в электронной таблице) еще не поддерживаются.
- Развертывание через портал администрирования Office 365 и AppSource еще не включено.
- Настраиваемые функции в Excel Online могут перестать работать во время сеанса после периода бездействия. Для восстановления функции обновите страницу веб-обозревателя (F5) и повторно введите настраиваемую функцию.
- Если у вас есть несколько надстроек, работающих на Excel для Windows, внутри ячейки таблицы может отображаться временный результат **#GETTING_DATA**. Закройте все окна Excel и перезапустите Excel.
- Возможно, в будущем появятся специальные средства отладки для настраиваемых функций. Тем временем вы можете выполнить отладку в Excel Online с помощью средств разработчика F12. Подробнее см. в статье [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md).

## <a name="changelog"></a>Журнал изменений

- **7 ноября 2017 г.**. Доставлена* предварительная версия настраиваемых функций с примерами
- **20 ноября 2017 года** исправлена ошибка совместимости для пользователей, использующих сборки 8801 и более новых версий
- **28 ноября 2017 г.**. Доставлена* поддержка отмены вызова асинхронных функций (необходимо изменение потоковых функций)
- **7 мая 2018 г.** Реализована*​​поддержка Mac, Excel Online и синхронных функций, выполняемых внутри процесса
- **20 сентября 2018 г.** Реализована поддержка среды выполнения JavaScript настраиваемых функций. Подробнее см. статью [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md).

\* на канале участников программы предварительной оценки Office

## <a name="see-also"></a>См. также

* [Метаданные настраиваемых функций](custom-functions-json.md)
* [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)
* [Руководство по настраиваемым функциям Excel](excel-tutorial-custom-functions.md)