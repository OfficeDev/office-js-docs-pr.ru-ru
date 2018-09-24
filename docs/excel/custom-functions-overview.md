---
ms.date: 09/20/2018
description: Создание настраиваемой функции в Excel с помощью JavaScript.
title: Создание настраиваемых функций в Excel (предварительная версия)
ms.openlocfilehash: 295152ca14cf56293d51b8b0512b729373841208
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062132"
---
# <a name="create-custom-functions-in-excel-preview"></a>Создание настраиваемых функций в Excel (предварительная версия)

Настраиваемые функции позволяют разработчикам добавлять новые функции в Excel, определяя эти функции в JavaScript как часть надстройки. Пользователи в Excel могут получать доступ к настраиваемым функциям, как к любой другой встроенной функции Excel (например, `SUM()`). В этой статье описано создание настраиваемых функций в Excel.

На следующем рисунке показан конечный пользователь, вставляющий пользовательскую функцию в ячейку листа Excel. Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которую пользователь указывает в качестве входных параметров для функции.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

Следующий код определяет настраиваемую функцию `ADD42`.

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

Настраиваемые функции теперь доступны в предварительной версии для разработчиков на Windows, Mac, а также в Excel Online. Чтобы попробовать их, выполните следующие действия.

1. Установите Office (сборка 10827 на Windows или 13.329 на Mac) и присоединитесь к программе [предварительной оценки Office](https://products.office.com/office-insider) . Вы должны присоединиться к программе предварительной оценки Office, чтобы иметь доступ к настраиваемым функциям. В настоящее время настраиваемые функции отключены во всех сборках Office, если вы не являетесь членом программы предварительной оценки Office.

2. Создайте проект надстройки настраиваемых функций Excel с помощью [Yo Office](https://github.com/OfficeDev/generator-office), а затем следуйте инструкциям в [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) для использования проекта.

3. Введите `=CONTOSO.ADD42(1,2)` в любой ячейке таблицы Excel, после чего нажмите клавишу **ВВОД**, чтобы запустить настраиваемую функцию.

> [!NOTE]
> В разделе [Известные проблемы](#known-issues) далее в этой статье указаны текущие ограничения настраиваемых функций.

## <a name="learn-the-basics"></a>Ознакомьтесь с основами

В проекте настраиваемых функций, который вы создали с помощью [Yo Office](https://github.com/OfficeDev/generator-office), вы увидите следующие файлы:

| Файл | Формат файла | Описание |
|------|-------------|-------------|
| **./src/customfunctions.js** | JavaScript | Содержит код, который определяет настраиваемые функции. |
| **./config/customfunctions.json** | JSON | Содержит метаданные, которые описывают настраиваемые функции и позволяют Excel регистрировать настраиваемые функции, чтобы сделать их доступными для пользователей. |
| **./index.html** | HTML | Предоставляет ссылку в тегах &lt;script&gt; на файл JavaScript, который определяет пользовательские функции. |
| **./manifest.xml** | XML | Указывает пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML, указанных ранее в этой таблице. |

### <a name="manifest-file-manifestxml"></a>Файл манифеста (./manifest.xml)

XML-файл манифеста для надстройки, который определяет настраиваемые функции, определяет пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML. Ниже показан пример использования элементов `<ExtensionPoint>` и `<Resources>` в разметке XML. Эти элементы необходимо включить в манифест надстройки, чтобы Excel мог выполнять настраиваемые функции.  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Функции Excel добавляются пространством имен, указанным в XML-файле манифеста. Пространство имен функции предшествует имени функции и отделяется от него точкой. Например, чтобы вызвать функцию `ADD42()` в ячейке листа Excel, следует ввести `=CONTOSO.ADD42`, так как CONTOSO — это пространство имен, а `ADD42` — имя функции, указанной в файле JSON. Пространство имен предназначено для использования в качестве идентификатора для вашей компании или надстройки. 

### <a name="json-file-configcustomfunctionsjson"></a>Файл JSON (. / config/customfunctions.json)

Файл метаданных настраиваемых функций предоставляет информацию, которую Excel требует для их регистрации, и делает их доступными для конечных пользователей. Настраиваемые функции регистрируются, когда пользователь в первый раз запускает надстройку. После этого пользователь может использовать их во всех книгах (то есть, не только в книге, в которой первоначально выполнялась надстройка).

> [!TIP]
> Чтобы настраиваемая функция работала корректно в Excel Online, в параметрах сервера, на котором размещается файл JSON, должен быть включен [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS).

Следующий код в файле **customfunctions.json** определяет метаданные для функции `ADD42`, описанной выше в этой статье. Эти метаданные определяют имя функции, ее описание, возвращаемое значение, входные параметры и многое другое. В таблице, следующей за этим примером кода, содержится подробная информация об отдельных свойствах этого объекта JSON.

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
                    "name": "number 1",
                    "description": "the first number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                },
                {
                    "name": "number 2",
                    "description": "the second number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
        }
    ]
}
```

В следующей таблице перечислены свойства, которые обычно присутствуют в файле метаданных JSON. Более подробные сведения о файле метаданных JSON, в том числе о параметрах, не использующихся в предыдущем примере, см. в статье [Метаданные настраиваемых функций](custom-functions-json.md).

| Свойство  | Описание |
|---------|---------|
| `id` | Уникальный идентификатор для функции. Этот идентификатор не должен изменяться после его установки. |
| `name` | Имя функции, отображаемое в меню автозаполнения, когда пользователь вводит формулу в ячейке. В меню автозаполнения это значение будет иметь префикс пространства имен настраиваемых функций, указанного в XML-файле манифеста. |
| `helpUrl` | URL-адрес страницы, которая отображается, когда пользователь запрашивает справку. |
| `description` | Описывает, что делает функция. Это значение появляется как подсказка, когда функция является выбранным элементом в меню автозаполнения в Excel. |
| `result`  | Объект, который определяет тип данных, который возвращается функцией. Значение дочернего свойства `type` может быть **string**, **number**или **boolean**. Значение дочернего свойства `dimensionality` может быть **scalar** или **matrix** (двухмерный массив значений указанного типа `type`). |
| `parameters` | Массив, который определяет входные параметры для функции. В Excel intelliSense появляются дочерние свойства `name` и `description`. Дочерние свойства `type` и `dimensionality` идентичны дочерним свойствам объекта `result`, описанного выше в этой таблице. |
| `options` | Это свойство позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию. Подробнее о том, как можно использовать это свойство, см. в разделах [Потоковые функции](#streamed-functions) и [Отмена](#canceling-a-function) ниже в этой статье. |

## <a name="functions-that-return-data-from-external-sources"></a>Функции, возвращающие данные из внешних источников

Если настраиваемая функция получает данные из внешнего источника, например веб-сайта, она должна:

1. возвращать обещание JavaScript в Excel;

2. разрешать Promise окончательным значением, используя функцию обратного вызова.

Настраиваемые функции отображают временный результат `#GETTING_DATA` в ячейке, когда Excel ожидает конечный результат. Во время ожидания результата пользователи могут нормально взаимодействовать с остальной частью листа.

В следующем примере кода настраиваемая функция `getTemperature()` получает от термометра текущую температуру. Обратите внимание, что функция `sendWebRequest` является гипотетической, не указанной здесь, и использует XHR для вызова веб-службы температуры.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>Потоковые функции

Потоковые настраиваемые функции позволяют вам выводить данные в ячейки многократно с течением времени, не требуя от пользователя явно запрашивать пересчет. Следующий пример кода — это настраиваемая функция, которая каждую секунду добавляет число к результату. Обратите внимание на следующие особенности этого кода:

- Excel автоматически отображает каждое новое значение при помощи `setResult` обратного вызова.

- Последний параметр, `handler`, никогда не указывается в коде регистрации и не отображается в меню автозаполнения, когда пользователи Excel вводят функцию. Это объект, который содержит функцию обратного вызова `setResult`, используемую для передачи данных из функции в Excel и обновления значения ячейки.

- Чтобы Excel передал функцию `setResult` объекту `handler`, необходимо объявить поддержку потоковой передачи при регистрации функции, установив параметр `"stream": true` для свойства `options` для настраиваемой функции в JSON-файле метаданных.

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a>Отмена функции

В некоторых случаях может потребоваться отменить выполнение потоковой настраиваемой функции, чтобы снизить ее потребление пропускной способности, рабочей памяти и загрузку процессора. Excel отменяет выполнение функции в следующих ситуациях.

- Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.

- Когда изменяется один из аргументов (входных параметров) функции. В этом случае после отмены активируется новый вызов функции.

- Пользователь вручную вызывает пересчет. В этом случае после отмены активируется новый вызов функции.

> [!NOTE]
> Вы должны реализовать обработчик отмены для каждой потоковой функции.

Чтобы сделать функцию отменяемой, установите для настраиваемой функции параметр `"cancelable": true` в свойстве `options` в JSON-файле метаданных.

В следующем коде показана та же функция `incrementValue`, которая была описана выше, но на этот раз с реализованным обработчиком отмены. В этом примере при отмене функции `incrementValue` будет выполняться метод `clearInterval()`.

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a>Сохранение и совместное использование состояния

Настраиваемые функции могут сохранять данные в глобальных переменных JavaScript. В последующих вызовах настраиваемая функция может использовать значения, сохраненные в этих переменных. Сохранение состояния может быть полезно, когда пользователи добавляют одну настраиваемую функцию к нескольким ячейкам, потому что все экземпляры функции могут совместно использовать ее состояние. Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.

В приведенном ниже коде показана реализация вышеописанной потоковой функции температуры, глобально сохраняющей состояние. Обратите внимание на следующие особенности этого кода:

- `refreshTemperature` — это потоковая функция, ежесекундно считывающая температуру определенного термометра. Новые температуры сохраняются в переменную `savedTemperatures`, но не обновляют значение ячейки напрямую. Она не должен вызываться непосредственно из ячейки листа, *поэтому она не регистрируется в файле JSON*.

- `streamTemperature` обновляет значения температуры, которые отображаются в ячейке каждую секунду, а в качестве источника данных использует переменную `savedTemperatures`. Она должна быть зарегистрирована в файле JSON и записана прописными буквами: `STREAMTEMPERATURE`.

- Пользователи могут вызывать функцию `streamTemperature` из нескольких ячеек в пользовательском интерфейсе Excel. Каждый вызов считывает данные из той же переменной `savedTemperatures`.

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

Например, предположим, что ваша функция возвращает второе наивысшее значение из диапазона чисел, хранящихся в Excel. Следующая функция принимает параметр `values`, который имеет тип `Excel.CustomFunctionDimensionality.matrix`. Обратите внимание, что в JSON-метаданных для этой функции вы должны для параметра `type` установить значение `matrix`.

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

При построении надстройки, определяющей настраиваемые функции, не забудьте добавить логику для обработки ошибок, возникающих в среде выполнения. Обработка ошибок для настраиваемых функций такая же, как и [обработка ошибок для Excel API JavaScript в целом](excel-add-ins-error-handling.md). В следующем примере кода метод `.catch` будет обрабатывать все ошибки, возникающие ранее в коде.

```js
function getComment(x) {
    //this delivers a section of lorem ipsum from the jsonplaceholder API
    let url = "https://jsonplaceholder.typicode.com/comments/" + x;

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
- Изменяемые функции (которые пересчитываются автоматически, когда в электронной таблице изменяются несвязанные данных) еще не поддерживаются.
- Развертывание через Портал администрирования Office 365 и AppSource еще не включено.
- Настраиваемые функции в Excel Online могут перестать работать во время сеанса после периода бездействия. Для восстановления работы обновите страницу браузера (F5) и повторно введите настраиваемую функцию.
- Если у вас есть несколько надстроек, работающих на Excel для Windows, внутри ячейки таблицы может отображаться временный результат **#GETTING_DATA**. Закройте все окна Excel и перезапустите Excel.
- Возможно, в будущем появятся специальные средства отладки для настраиваемых функций. Тем временем вы можете выполнить отладку в Excel Online с помощью средств разработчика F12. Подробнее см. в статье [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md).

## <a name="changelog"></a>Журнал изменений

- **7 ноября 2017 г.**. Выпущена* предварительная версия настраиваемых функций с примерами
- **20 ноября 2017 г.** Исправлена ошибка совместимости для пользователей, использующих сборки 8801 и выше.
- **28 ноября 2017 г.**. Выпущена* поддержка отмены вызова асинхронных функций (необходимо изменение для потоковых функций)
- **7 мая 2018 г.**. Выпущена*​​поддержка Mac, Excel Online и синхронных функций, выполняемых внутри процесса
- **20 сентября 2018 г.**. Выпущена поддержка среды выполнения JavaScript настраиваемых функций. Подробнее см. статью [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md).

\* канал участников программы предварительной оценки Office

## <a name="see-also"></a>См. также

* [Метаданные настраиваемых функций](custom-functions-json.md)
* [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)