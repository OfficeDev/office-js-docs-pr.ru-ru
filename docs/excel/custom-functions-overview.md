# <a name="create-custom-functions-in-excel-preview"></a>Создание настраиваемых функций в Excel (ознакомительная версия)

Настраиваемые функции (подобные пользовательским функциям, или UDF) позволяют разработчикам добавить любую функцию JavaScript в Excel с помощью надстройки. После этого пользователи смогут получать доступ к настраиваемым функциям, как к любой другой встроенной функции Excel (например, "=СУММ()"). В этой статье описано создание настраиваемых функций в Excel.

На приведенном ниже рисунке показано, как работают настраиваемые функции в пользовательском интерфейсе Excel.

<img src="../images/custom-function.gif" width="579" height="383" />

Ниже представлен пример кода настраиваемой функции, которая прибавляет 42 к паре чисел.

```js
function add42 (a, b) {
    return a + b + 42;
}
```

Настраиваемые функции уже доступны в ознакомительной версии. Чтобы опробовать их, выполните указанные ниже действия.

1.  Установите Office 2016 для Windows и присоединитесь к [программе предварительной оценки Office](https://products.office.com/ru-RU/office-insider).
2.  Клонируйте репозиторий [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) и следуйте указаниям из файла *README.md*, чтобы запустить надстройку в Excel.
3.  Введите `=CONTOSO.ADD42(1,2)` в любой ячейке и нажмите клавишу **ВВОД**, чтобы выполнить настраиваемую функцию.

В разделе "Известные проблемы" в конце этой статьи указаны текущие ограничения на настраиваемые функции, которые со временем будут обновляться.

## <a name="learn-the-basics"></a>Основы


В клонированном примере репозитория вы увидите перечисленные ниже файлы.

-   Файл *customfunctions.js*, который содержит следующее:

    -   Код настраиваемой функции, добавляемый в Excel.
    -   Код регистрации для подключения настраиваемой функции к Excel. После регистрации настраиваемые функции отображаются в списке доступных функций, появляющемся при вводе текста в ячейках.
-   Файл *customfunctions.html*, который содержит ссылку &lt;Script&gt; на файл *customfunctions.js*. Этот файл не отображается в пользовательском интерфейсе Excel.
-   Файл *manifest.xml*, который сообщает приложению Excel расположение HTML- и JS-файлов, необходимых для выполнения настраиваемых функций.

### <a name="javascript-file-customfunctionsjs"></a>Файл JavaScript (*customfunctions.js*)

Приведенный ниже код из файла customfunctions.js объявляет настраиваемую функцию `add42`, а затем регистрирует ее в Excel.

```js
function add42 (a, b) {
    return a + b + 42;
}

Excel.Script.customFunctions["CONTOSO"]["ADD42"] = {
    call: add42,
    description: "Adds 42 to the sum of two numbers",
    helpUrl: "https://www.contoso.com/help.html",
    result: {
        resultType: Excel.CustomFunctionValueType.number,
        resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    parameters: [{
        name: "num 1",
        description: "The first number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    {
        name: "num 2",
        description: "The second number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    }],
    options:{ batch: false, stream: false }
};

Excel.run(function(ctx) {
    ctx.workbook.customFunctions.addAll();
});
```

Для **регистрации** настраиваемой функции используется блок кода `Excel.Script.customFunctions["CONTOSO"]["ADD42"]`. Для регистрации функции в Excel необходимы указанные ниже параметры.

-   Префикс и имя функции: первое значение в `Excel.Script.customFunctions` — это префикс (в данном случае указан префикс CONTOSO); второе значение в `Excel.Script.customFunctions` — это имя функции (в данном случае указано имя ADD42). В Excel префикс и имя функции разделены точкой. Чтобы использовать настраиваемую функцию, объедините префикс функции (CONTOSO) с ее именем (ADD42) и введите `=CONTOSO.ADD42` в ячейке. По соглашению префиксы и имена функций указываются прописными буквами. Префикс служит в качестве идентификатора надстройки.
-   `call`. Определяет вызываемую функцию JavaScript (например, `add42`). Имя функции JavaScript может не совпадать с именем, зарегистрированным в Excel.
-   `description`. Описание отображается в меню автозаполнения в Excel.
-   `helpUrl`. Когда пользователь запрашивает справку по функции, Excel открывает область задач, в которой отображается веб-страница, расположенная по этому URL-адресу.
-   `result`. Определяет тип данных, возвращаемых функцией в Excel.

    -   `resultType`. Функция может возвращать значения типа `"string"` или `"number"` (также используется для дат и денежных сумм). Дополнительные сведения см. в статье [Перечисления настраиваемых функций](http://dev.office.com/reference/add-ins/excel/customfunctionsenumerations).
    -   `resultDimensionality`. Функция может возвращать одно значение (`"scalar"`) или `"matrix"` (матрицу значений). При возвращении матрицы значений функция возвращает массив, каждый элемент которого является массивом, представляющим строку значений. Дополнительные сведения см. в статье [Перечисления настраиваемых функций](http://dev.office.com/reference/add-ins/excel/customfunctionsenumerations). В приведенном ниже примере возвращается матрица из 3 строк и 2 столбцов со значениями из настраиваемой функции.

```js
return [["first","row"],["second","row"],["third","row"]];
```

-   Настраиваемая функция может принимать аргументы в качестве входных данных. Аргументы, передаваемые настраиваемой функции, указываются в свойстве *parameters*. Порядок параметров в определении должен соответствовать их порядку в функции JavaScript. Для каждого параметра определите указанные ниже свойства.

    -   `name`. Строка, представляющая параметр в Excel.
    -   `description`. Строка с дополнительными сведениями о параметре.
    -   `valueType`. Значение `"number"` или `"string"` по аналогии с вышеописанным свойством resultType.
    -   `valueDimensionality`. Значение `"scalar"` или `"matrix"` (матрица значений) по аналогии с вышеописанным свойством resultDimensionality. С помощью параметров матричного типа пользователи могут выбирать диапазоны из нескольких ячеек.

-   `options` позволяет использовать настраиваемые функции специальных типов, которые подробнее рассматриваются далее в этой статье.

Чтобы завершить регистрацию всех функций, определенных с помощью `Excel.Script.customFunctions`, обязательно вызовите метод `CustomFunctions.addAll()`.

После регистрации настраиваемые функции становятся доступны пользователю во всех книгах (а не только в той, где надстройка работала изначально). Функции отображаются в меню автозаполнения, когда пользователь начинает вводить название.

### <a name="manifest-file-manifestxml"></a>Файл манифеста (*manifest.xml*)

В приведенном ниже примере файла manifest.xml приложению Excel разрешается находить код для функций.

```xml

<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">

    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="scriptURL" />
                        <!— Required. The Developer Preview does not use the Script element.-->
                    </Script>
                    <Page>
                        <SourceLocation resid="pageURL"/>
                    </Page>
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>

    <Resources>
        <bt:Urls>
            <bt:Url id="scriptURL" DefaultValue="https://www.contoso.com/addin/customfunctions.js" />
            <bt:Url id="pageURL" DefaultValue="https://www.contoso.com/addin/customfunctions.html" />
        </bt:Urls>
    </Resources>

</VersionOverrides>

```

В приведенном выше коде задано следующее:

-   Элемент &lt;`Script`&gt;, являющийся обязательным, но не используемый в версии Developer Preview.
-   Элемент &lt;`Page`&gt;, ссылающийся на HTML-страницу надстройки. HTML-страница включает ссылку &lt;Script&gt; на файл JavaScript (*customfunctions.js*), содержащий настраиваемую функцию и код регистрации. HTML-страница скрыта и никогда не отображается в пользовательском интерфейсе.

## <a name="asynchronous-functions"></a>Асинхронные функции

Если настраиваемая функция получает данные из Интернета, необходимо выполнить асинхронный вызов, чтобы получить ее. При вызове внешних веб-служб настраиваемая функция должна:

1.   Возвращать обещание JavaScript в Excel.
2.   Отправлять HTTP-запрос на вызов внешней службы.
3.   Разрешать обещание с помощью метода обратного вызова `setResult`. Метод `setResult` отправляет значение в Excel.

В приведенном ниже коде показан пример настраиваемой функции, получающей температуру термометра.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult, setError){
        sendWebRequestExample(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>Потоковые функции

С помощью потоковых настраиваемых функций вы можете многократно выводить данные в ячейки, не дожидаясь, пока Excel или пользователь запросит повторное вычисление. Например, настраиваемая функция `incrementValue` в приведенном ниже коде ежесекундно прибавляет число к результату, а каждое новое значение отображается в Excel с помощью метода обратного вызова `setResult`. Пример использования кода регистрации с `incrementValue` вы найдете в файле *customfunctions.js*.

```js
function incrementValue(increment, setResult){ 
     var result = 0;
     setInterval(function(){
         result += increment;
         setResult(result);
    }, 1000);
}
```

Для потоковых функций последний параметр, `setResult`, никогда не указывается в коде регистрации и не отображается в меню автозаполнения, когда пользователи Excel вводят функцию. Это функция обратного вызова, используемая для передачи данных из функции в Excel и обновления значения ячейки. Чтобы приложение Excel могло передать функцию `setResult`, необходимо объявить поддержку потоковой передачи при регистрации функции, задав для параметра `stream` значение `true`.

## <a name="saving-state"></a>Сохранение состояния

Настраиваемые функции могут сохранять данные в глобальных переменных JavaScript. В последующих вызовах настраиваемая функция может использовать значения, сохраненные в этих переменных. Сохраненное состояние удобно использовать, если пользователи вводят несколько экземпляров одной функции, которые должны совместно использовать данные. Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.

В приведенном ниже коде показана реализация вышеописанной функции передачи температуры, сохраняющей состояние с помощью переменной `savedTemperatures`. В этом коде демонстрируются следующие понятия:

-   **Сохранение данных.** `refreshTemperature` — это потоковая функция, ежесекундно считывающая температуру определенного термометра. Новые значения температуры сохраняются в переменной savedTemperatures.

-   **Использование сохраненных данных.** Функция `streamTemperature` ежесекундно обновляет значения температуры, отображаемые в пользовательском интерфейсе Excel. Температуры считываются из переменной `savedTemperature`, а затем отправляются в пользовательский интерфейс Excel с помощью метода `setResult`. Пользователи могут вызывать функцию `streamTemperature` из нескольких ячеек в пользовательском интерфейсе Excel. При каждом вызове функции `streamTemperature` считываются данные из переменной `savedTemperatures`.

> В этом случае мы регистрируем `streamTemperature` как настраиваемую функцию в Excel.

```js
var savedTemperatures{};

function streamTemperature(thermometerID, setResult){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID);
     }

     function getNextTemperature(){
         setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequestExample(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a>Работа с диапазонами данных

Настраиваемая функция может принимать диапазон данных в качестве параметра или возвращать диапазон данных.

Допустим, функция возвращает вторую по величине температуру из диапазона значений, хранящихся в Excel. Приведенная ниже функция принимает параметр `temperatures`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.

```js
function secondHighestTemp(temperatures){ 
     var highest = -273, secondHighest = -273;
     for(var i = 0; i < temperatures.length; i++){
         for(var j = 0; j < temperatures[i].length; j++){
             if(temperatures[i][j] <= highest){
                 secondHighest = highest;
                 highest = temperatures[i][j];
             }
             else if(temperatures[i][j] <= secondHighest){
                 secondHighest = temperatures[i][j];
             }
         }
     }
     return secondHighest;
 }
```

## <a name="known-issues"></a>Известные проблемы

Указанные ниже функции еще не поддерживаются в версии Developer Preview.

-   Пакетная обработка, позволяющая агрегировать несколько вызовов одной функции для повышения производительности.

-   Отмена с уведомлением о том, что потоковая функция больше не требуется (например, когда пользователь очищает ячейку). В настоящее время функции не могут определять, когда следует перестать записывать в ячейку новые значения.

-   URL-адреса справки и описания параметров пока не используются в Excel.

-   Публикация надстроек, использующих настраиваемые функции, в Магазине Office и их централизованное развертывание в Office 365.

-   Настраиваемые функции недоступны в Excel для Mac, Excel для iOS и Excel Online.

-   В настоящее время надстройки используют скрытый процесс браузера для выполнения настраиваемых функций. В будущем JavaScript будет работать на некоторых платформах напрямую, чтобы настраиваемые функции выполнялись быстрее и использовали меньше памяти. Кроме того, HTML-страница, на которую ссылается элемент &lt;Page&gt; манифеста, не будет необходима для большинства платформ, так как Excel будет выполнять код JavaScript напрямую. Чтобы подготовиться к этому изменению, убедитесь, что в ваших настраиваемых функциях не используется модель DOM для веб-страниц.
