---
ms.date: 01/30/2019
description: Создание пользовательских функций в Excel с помощью JavaScript.
title: Создание пользовательских функций в Excel (ознакомительная версия)
localization_priority: Priority
ms.openlocfilehash: 312a590052f1f78c8ff5477c8cfb85eb94f03aad
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052765"
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

Например, приведенный ниже код определяет пользовательские функции `add` и `increment`, а затем указывают информацию о сопоставлении для обеих функций. Функция `add` сопоставляется с объектом в файле метаданных JSON, где значение свойства `id` **ADD**, и функция `increment` будет сопоставляться с объектом в файле метаданных, где значение свойства`id` **INCREMENT**. См. статью [Советы и рекомендации по работе с пользовательскими функциями](custom-functions-best-practices.md#associating-function-names-with-json-metadata) для получения дополнительных данных о сопоставлении имен функций в файле скрипта с объектами в файле метаданных JSON.

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
 CustomFunctions.associate("INCREMENT", increment);
```

### <a name="json-metadata-file"></a>Файл метаданных JSON

Файл метаданных пользовательских функций (**./config/customfunctions.json** в проекте, созданном во время генератора Yo Office) предоставляет информацию, которая необходима Excel для регистрации пользовательских функций и обеспечения их доступности для конечных пользователей. Пользовательские функции регистрируются, когда пользователь запускает надстройку в первый раз. После этого как они становятся доступны тому самому пользователю во всех рабочих книгах (т.е. не только в рабочей книге, где надстройка первоначально запущена).

> [!TIP]
> Настройки сервера на сервере, на котором размещен JSON-файл, должны включать активацию [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS), чтобы пользовательские функции сработали надлежащим образом в Excel Online.

Код ниже в **customfunctions.json** определяет метаданные для функции `add` и функции `increment`, описанные ранее. Таблица, которая следует за этим примером кода, предоставляет подробные сведения об отдельных свойств для этого объекта JSON. См. статью [Советы и рекомендации по работе с пользовательскими функциями](custom-functions-best-practices.md#associating-function-names-with-json-metadata) для получения дополнительных данных об указании имен свойств `id` и `name` в файле метаданных JSON.

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
| `options` | Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию. Дополнительные сведения о способах использования этого свойства см. в разделах [Потоковая передача функций](#streaming-functions) и [Отмена функции](#canceling-a-function). |

### <a name="manifest-file"></a>Файл манифеста

XML-файл манифеста для надстройки, который определяет пользовательские функции (**./manifest.xml** в проекте, который создает генератор Yo Office) и определяет пространство имен для всех пользовательских функций в надстройке, а также расположение файлов JavaScript, JSON и HTML. XML-разметка ниже представляет пример элементов `<ExtensionPoint>` и `<Resources>`, которые необходимо включить в манифест надстройки, чтобы активировать пользовательские функции.  

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://localhost:8081/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://localhost:8081/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://localhost:8081/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
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

## <a name="declaring-a-volatile-function"></a>Объявление переменной функции

[Переменные функции](https://docs.microsoft.com/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) — это функции, значение которых периодически изменяется, даже если никакой из аргументов функции не меняется. Эти функции пересчитываются при каждом пересчете в Excel. К примеру, представьте себе ячейку, вызывающую функцию `NOW`. При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.

В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`. Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](https://docs.microsoft.com/ru-RU/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Пользовательские функции позволяют создавать собственные переменные функции, которые могут быть полезны при обработке дат, времени, случайных чисел и моделировании. Например, при моделированиях методом Монте-Карло требуется создание случайных входных данных, чтобы определить оптимальное решение.

Чтобы объявить функцию переменной, добавьте `"volatile": true` в объект `options` для функции в файле метаданных JSON, как показано в приведенном ниже примере кода. Обратите внимание, что функция не может одновременно иметь значения `"streaming": true` и `"volatile": true`. Если оба параметра помечены как `true`, параметр переменности будет игнорироваться.

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

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

## <a name="co-authoring"></a>Совместное редактирование
Excel Online и Excel для Windows с подпиской на Office 365 позволяют совместно редактировать документы. Эта функция работает с пользовательскими функциями. Если в книге используется пользовательская функция, вашему коллеге будет предложено загрузить надстройку пользовательской функции. Когда вы оба загрузите надстройку, пользовательская функция поделится результатами с помощью совместного редактирования.

Дополнительные сведения о совместном редактировании см. в статье [О совместном редактировании в Excel](https://docs.microsoft.com/ru-RU/office/vba/excel/concepts/about-coauthoring-in-excel).

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

## <a name="determine-which-cell-invoked-your-custom-function"></a>Определение того, какая ячейка вызывала пользовательскую функцию

В некоторых случаях вам потребуется получить адрес ячейки, которая вызывала пользовательскую функцию. Это может быть полезно в следующих типах сценариев:

- Форматирование диапазонов: Используйте адрес ячейки в качестве ключа для хранения сведений в [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). После этого используйте событие [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) в Excel, чтобы загрузить ключ из `AsyncStorage`.
- Отображение кэшированных значений. Если функция используется в автономном режиме, отображайте сохраненные в кэше значения из `AsyncStorage` с помощью `onCalculated`.
- Сверка: используйте адрес ячейки, чтобы найти исходную ячейку, чтобы упростить сверку при выполнении обработки.

Сведения об адресе ячейки предоставляются только в том случае, если параметру `requiresAddress` присвоено значение `true` в файле метаданных JSON функции. Ниже приведен пример:

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

Чтобы найти адрес ячейки, в файл скрипта (**./src/customfunctions.js** или **./src/customfunctions.ts**) потребуется также добавить функцию `getAddress`. В этой функции можно использовать параметры, как показано в примере ниже в виде `parameter1`. В качестве последнего параметра всегда будет использоваться `invocationContext` — объект, содержащий расположение ячейки, которое передает приложение Excel, если параметру `requiresAddress` присвоено значение `true` в файле метаданных JSON.

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

По умолчанию значения, возвращаемые из функции `getAddress`, соответствуют следующему формату: `SheetName!CellNumber`. Например, если функция вызвана с листа с названием Expenses (Расходы) в ячейке B2, возвращаемым значением будет `Expenses!B2`.

## <a name="handling-errors"></a>Обработка ошибок

При создании надстройки, которая определяет пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Обработка ошибок для пользовательских функций в значительной степени совпадает с [обработкой ошибок для API JavaScript в Excel](excel-add-ins-error-handling.md). В следующем примере кода `.catch` будет обрабатывать любые ошибки, возникающие ранее в коде.

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

С известными проблемами можно ознакомиться в нашем [репозитории GitHub, посвященном пользовательским функциям в Excel](https://github.com/OfficeDev/Excel-Custom-Functions/issues). 

## <a name="see-also"></a>См. также

* [Метаданные пользовательских функций](custom-functions-json.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Рекомендации по пользовательским функциям](custom-functions-best-practices.md)
* [Журнал изменений пользовательских функций](custom-functions-changelog.md)
* [Руководство по настраиваемым функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)

