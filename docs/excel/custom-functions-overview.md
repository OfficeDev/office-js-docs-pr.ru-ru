---
ms.date: 10/17/2018
description: Создание пользовательских функций в Excel с помощью JavaScript.
title: Создание пользовательских функций в Excel (Ознакомительная версия)
ms.openlocfilehash: 8383b5f6d568a1ce2da036fbacfb90404bbe8297
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298553"
---
# <a name="create-custom-functions-in-excel-preview"></a>Создание пользовательских функций в Excel (ознакомительная версия)

Пользовательские функции позволяют разработчикам добавлять новые функции в Excel, посредством определения этих функций в JavaScript как части надстройки. Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`. В этой статье описано создание специальных функций в Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Ниже продемонстрировано, как конечный пользователь, вставляет настраиваемую функцию в ячейке на листе Excel. Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которые пользователь указывает в качестве входных параметров для функции.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

Приведенный ниже код определяет настраиваемую функцию `ADD42`.

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> В разделе [Известные проблемы](#known-issues) далее в этой статье определены текущие ограничения для пользовательских функций.

## <a name="components-of-a-custom-functions-add-in-project"></a>Компоненты пользовательские функции для надстройки проекта.

Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания в Excel проекта с пользовательскими функциями, вы увидите следующие файлы в проекте, созданном генератором:

| Файл | Формат файла | Описание |
|------|-------------|-------------|
| **./src/customfunctions.js**<br/>или<br/>**./src/customfunctions.ts** | JavaScript<br/>или<br/>TypeScript | Содержит код, который определяет пользовательские функции. |
| **./config/customfunctions.json** | JSON | Содержит метаданные с описанием пользовательских функций и позволяет Excel регистрировать пользовательские функции и сделать их доступными для конечных пользователей. |
| **./index.html** | HTML | Предоставляет &lt;скрипт&gt; со ссылкой на файл JavaScript, который определяет пользовательские функции. |
| **./manifest.xml** | XML | Определяет пространство имен для всех пользовательских функций в надстройку и расположение JavaScript, JSON и HTML-файлов, которые указаны ранее в этой таблице. |

В разделах ниже приведены дополнительные сведения о данных файлах.

### <a name="script-file"></a>Файл скрипта 

Файл сценария (**./src/customfunctions.js** или **./src/customfunctions.ts** в проекте, созданном генератором Yo Office) содержит код, который определяет пользовательские функции и размещает имена пользовательских функций к объектам в [файле метаданных JSON](#json-metadata-file). 

Например, приведенный ниже код определяет пользовательские функции `add` и `increment`, а затем указывают информация о сопоставлении для обоих функций. Функция `add` будет сопоставлена с объектом в файле метаданных JSON, где значение свойства `id` **ADD**, и функция `increment` будет сопоставлена с объектом в файле метаданных, где значение свойства`id` **INCREMENT**. См. статью [Советы и рекомендации по работе с пользовательскими функциями](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) для получения дополнительных данных о сопоставление имен функций в файле скрипта с объектами в файле метаданных JSON.

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

Файл метаданных пользовательских функций (**./config/customfunctions.json** в проекте, созданном во время генератора Yo Office) предоставляет информацию, которая необходима Excel для регистрации пользовательских функций и обеспечения их доступности для конечных пользователей. Пользовательские функции регистрируются, когда пользователь запускает надстройку в первый раз. После этого как они становятся доступны тому самому пользователю во всех рабочих книгах (т.е. не только в рабочей книге, где надстройка первоначально запущена).

> [!TIP]
> Настройки сервера на сервере, на котором размещен JSON-файл, должны включать активацию [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS), чтобы пользовательские функции сработали надлежащим образом в Excel Online.

Код ниже в **customfunctions.json** определяет метаданные для функции `add` и функции `increment`, описанные ранее. Таблица, которая следует за этим примером кода, предоставляет подробные сведения об отдельных свойств для этого объекта JSON. См. статью [Советы и рекомендации по работе с пользовательскими функциями](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) для получения дополнительных данных об указании имен свойств `id` и `name` в файле метаданных JSON.

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

В таблице ниже перечислены свойства, которые обычно есть в файле метаданных JSON. Дополнительные сведения о файле метаданных JSON см. в статье [Пользовательские функции метаданных](custom-functions-json.md).

| Свойство  | Описание |
|---------|---------|
| `id` | Уникальный идентификатор для функции. Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки. |
| `name` | Имя функции, которая будет отображаться пользователю в Excel. В Excel это имя функции будет включать префикс пространства имен пользовательских функций, который указан в [XML файле манифеста](#manifest-file). |
| `helpUrl` | URL-адрес страницы, который отображается при запросе пользователем справки. |
| `description` | Описание того, что делает функция. Это значение отображается в виде подсказки, когда функция представляет собой выделенный элемент в меню автозаполнения в Excel. |
| `result`  | Объект, который определяет тип информации, возвращаемый функцией. Для получения более подробной информации об этом объекте см. [результат](custom-functions-json.md#result). |
| `parameters` | Массив, который определяет входные параметры для функции. Для получения более подробной информации об этом объекте см. [параметры](custom-functions-json.md#parameters). |
| `options` | Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию. Дополнительные сведения о способах использования этого свойства см. разделы [Потоковая передача функции](#streaming-functions) и [Отмена функция](#canceling-a-function) ниже в этой статье. |

### <a name="manifest-file"></a>Файл манифеста

XML-файл манифеста для надстройки, который определяет пользовательские функции (**./manifest.xml** в проекте, который создает генератор Yo Office) и определяет пространство имен для всех пользовательских функций в надстройке, а также расположение файлов JavaScript, JSON и HTML. XML-разметка ниже представляет пример элементов `<ExtensionPoint>` и `<Resources>`, которые необходимо включить в манифест надстройки, чтобы активировать пользовательские функции.  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. Can only contain alphanumeric characters and periods.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Функции в Excel имеют в начале пространство имен, указанное в XML-файле манифеста. Пространство имен функции предшествует названию функции, и они будут разделены точкой. Например, чтобы вызвать функцию `ADD42` в ячейке на листе Excel, введите `=CONTOSO.ADD42`, так как `CONTOSO` является пространством имен, а `ADD42` — это имя функции, определяемой в JSON-файл. Пространство имен служит в качестве идентификатора для вашей компании или надстройки. Пространство имен может содержать только буквы, цифры и точки.

## <a name="functions-that-return-data-from-external-sources"></a>Функции, которые возвращают данные из внешних источников

Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:

1. Возвращать обещание JavaScript в Excel;

2. Устранять обещание с итоговым значением с помощью функции обратного вызова.

Пользовательские функции отображают `#GETTING_DATA` временный результат в ячейке, пока Excel ожидает конечный результат. Пользователи могут нормально взаимодействовать с остальным листом, хотя они ожидают результат.

В приведенном ниже примере кода пользовательская функция `getTemperature()` возвращает текущую температуру термометра. Обратите внимание, что `sendWebRequest` — это гипотетическая функция (не указанная ниже), которая использует [XHR](custom-functions-runtime.md#xhr-example) для вызова веб-службы.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a>Потоковая передача функций

Потоковая передача пользовательских функций позволяет выводить данные в ячейки несколько раз в течением времени, избавляя пользователя от необходимости явным образом запрашивать обновление данных. Приведенный ниже пример кода — это настраиваемая функция, которая добавляет число к результату каждую секунду. Обратите внимание на следующие особенности этого кода:

- Excel отображает каждое новое значением автоматически с помощью обратного вызова `setResult`.

- Второй параметр ввода, `handler`, не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".

- Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции. Вам необходимо реализовать уведомление об отмене следующим образом для любой функции потоковой передачи. Дополнительные сведения см. в статье [Отмена функции](#canceling-a-function).

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

Когда вы указываете метаданные для функции потоковой передачи в файле метаданных JSON, необходимо задать свойства `"cancelable": true` и `"stream": true` в объекте `options`, как показано в приведенном ниже примере.

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

В некоторых случаях может потребоваться отмена выполнения пользовательских функций потоковой передачи, чтобы уменьшить использования пропускной способности, рабочей памяти и загрузку ЦП. Excel отменяет выполнение функций в следующих случаях:

- Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.

- Когда изменяется один из аргументов (входных параметров) функции. В этом случае после отмены выполняется новый вызов функции.

- Когда пользователь вручную вызывает пересчет. В этом случае после отмены выполняется новый вызов функции.

Чтобы активировать возможность отмены функции, необходимо реализовать обработчик отмены в функции JavaScript, а также указать свойство `"cancelable": true` в объекте `options` в метаданных JSON, который описывает функцию. Примеры кода в предыдущем разделе этой статьи предоставляют собой пример использования данных техник.

## <a name="saving-and-sharing-state"></a>Состояние сохранения и совместного использования

Пользовательские функции могут сохранять данные в глобальных переменных JavaScript, которые можно использовать в последующих вызовах. Сохраненное состояние полезно, когда пользователи вызывают одни и те же настраиваемые функций из более чем одной ячейки, так как все экземпляры функции могут получить доступ к состоянию. Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.

В приведенном ниже примере кода показана реализация вышеописанной функции передачи температуры, сохраняющей состояние с помощью глобальной переменной. Обратите внимание на следующие особенности этого кода:

- Функция `streamTemperature` обновляет значение температуры, которое отображается в ячейке, каждую секунду и использует переменную `savedTemperatures` как источник данных.

- Так как `streamTemperature` — это функция потоковой передачи, она реализует обработчик отмены, который будет запускаться, если функция была отменена.

- Если пользователь вызывает функцию `streamTemperature` из нескольких ячеек в Excel, функция `streamTemperature` считывает данные из той же самой переменной `savedTemperatures` при каждом запуске. 

- Функция `refreshTemperature` ежесекундно считывает температуру определенного термометра и сохраняет результат в переменной `savedTemperatures`. Так как функция `refreshTemperature` недоступна для конечных пользователей в Excel, ее не нужно регистрировать в JSON-файле.

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
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

Ваша пользовательская функция может принимать широкий диапазон данных в виде входных параметров или возвращать широкий диапазон данных. В JavaScript диапазон данных будет иметь вид двумерного массива.

Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel. Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`. Обратите внимание, что в метаданных JSON для данной функции вам следует задать для параметра свойство `type` в `matrix`.

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

При создании надстройки, которая определяет пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Обработка ошибок для пользовательских функций совпадает с [обработкой ошибок для Excel JavaScript API ошибок в значительной степени](excel-add-ins-error-handling.md). В следующем примере кода `.catch` будет обрабатывать любые ошибки, возникающие ранее в коде.

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

- URL-адреса справки и описания параметров в настоящее время не используются Excel.
- Пользовательские функции в настоящее время недоступны в Excel для мобильных клиентов.
- Переменные функции (которые пересчитываются автоматически всякий раз при изменениях несвязанных данных на листе) еще не поддерживаются.
- Развертывание через портал администрирования Office 365 и AppSource еще не активировано.
- Пользовательские функции в Excel Online могут перестать работать во время сеанса после периода бездействия. Обновите страницу браузера (F5) и еще раз введите пользовательскую функции для восстановления работоспособности.
- Вы можете увидеть временный результат **#GETTING_DATA** (# ОЖИДАНИЕ_ДАННЫХ) внутри ячейки(-ек), листа, если у вас есть несколько надстроек, запущенных в Excel для Windows. Закройте все окна Excel и перезапустите Excel.
- Инструменты для отладки, предназначенные специально для пользовательских функций, могут быть доступны в будущем. В настоящее время вы можете выполнить отладку в Excel Online при использовании средств разработчика F12. Дополнительные данные см. [Советы и рекомендации в отношении пользовательских функций](custom-functions-best-practices.md)

## <a name="changelog"></a>Журнал изменений

- **7 ноября 2017 г.**: Выпущена ознакомительная версия пользовательских функций с примерами.
- **20 ноября 2017 г.**: Исправлена ошибка совместимости для пользователей, использующих сборки 8801 и выше.
- **28 ноября 2017 г.**: Добавлена поддержка отмены вызова асинхронных функций (необходимо изменение для потоковых функций).
- **7 мая 2018 г.**: Реализована* поддержка запущенный подпроцессов для Mac, Excel Online и синхронных функций
- **20 сентября 2018 г.**: Реализована поддержка пользовательских функций среды выполнения JavaScript. Дополнительные сведения см. в статье [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md).
- **20 октября 2018 г.**: После выхода [Сборки October Insiders](https://support.office.com/ru-RU/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), пользовательские функции требуют параметр «идентификатор» в [метаданных пользовательских функций](custom-functions-json.md) для настольных версий Windows и Online. На компьютерах Mac можно игнорировать этот параметр.


\* к каналу [Office Insider ](https://products.office.com/office-insider) (ранее "Предварительная оценка — ранний доступ")

## <a name="see-also"></a>См. также

* [Метаданные пользовательских функций](custom-functions-json.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Советы и рекомендации в отношении пользовательских функций](custom-functions-best-practices.md)
* [Руководство по пользовательским функциям в Excel](excel-tutorial-custom-functions.md)
