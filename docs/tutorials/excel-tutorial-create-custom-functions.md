---
title: Руководство по настраиваемым функциям в Excel (предварительная версия)
description: Из этого руководства вы узнаете, как создать надстройку, Excel, содержащую пользовательские функции, которые могут выполнять вычисления, запрашивать или передавать веб-данные.
ms.date: 01/08/2019
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 9491b29094eb486f7efbe7e128a7a77be43bff39
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635953"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a>Руководство: создание пользовательских функций в Excel (предварительная версия)

Пользовательские функции позволяют добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки. Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`. Вы можете создавать пользовательские функции, которые будут выполнять простые задачи, такие как вычисления, или более сложные задачи, такие как потоковая передача данных в режиме реального времени из Интернета на лист.

В этом руководстве описан порядок выполнения перечисленных ниже задач.
> [!div class="checklist"]
> * Создание надстройки пользовательской функции с помощью [генератора Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office). 
> * Использование готовой пользовательской функции для выполнения простых вычислений
> * Создание пользовательской функции, которая получает данные из сети Интернет.
> * Создание пользовательской функции, которая осуществляет потоковую передачу данных в реальном времени из сети Интернет

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>Необходимые компоненты

* [Node.js](https://nodejs.org/en/) (версия 8.0.0 или более поздняя)

* [Git Bash](https://git-scm.com/downloads) (или другой клиент Git)

* Последняя версия [Yeoman](https://yeoman.io/) и [генератора Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Даже если у вас установлен генератор Yeoman, рекомендуется обновить пакет до последней версии из npm.

* Excel для Windows (64-разрядная версия 1810 или более поздняя) или Excel Online

* Присоединитесь к [Программе предварительной оценки Office](https://products.office.com/office-insider) (уровень **Участник**; ранее "Предварительная оценка — ранний доступ")

## <a name="create-a-custom-functions-project"></a>Создание проекта пользовательских функций

 Чтобы начать, вам необходимо создать проект кода для разработки надстройки пользовательской функции. [Генератор Yeoman для надстройки Office](https://www.npmjs.com/package/generator-office) настроит ваш проект с некоторыми начальными пользовательскими функциями, которые вы можете попробовать использовать.

1. Выполните указанную ниже команду и ответьте на вопросы, как показано ниже.
    
    ```
    yo office
    ```
    
    * Выберите тип проекта: `Excel Custom Functions Add-in project (...)`
    * Выберите тип сценария: `JavaScript`
    * Как вы хотите назвать свою надстройку? `stock-ticker`
    
    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/12-10-fork-cf-pic.jpg)
    
    Генератор Yeoman создает файлы проекта и устанавливает вспомогательные компоненты Node.js.

2. Перейдите в папку проекта.
    
    ```
    cd stock-ticker
    ```

3. Сделайте доверенным самозаверяющий сертификат, необходимый для выполнения этого проекта. Подробные инструкции для Windows или Mac см. в статье [Добавление самозаверяющих сертификатов в качестве доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Выполните сборку проекта.
    
    ```
    npm run build
    ```

5. Запустите локальный веб-сервер, работающий на Node.js. Вы можете попробовать использовать надстройку пользовательской функции в Excel для Windows или в Excel Online.

# <a name="excel-for-windowstabexcel-windows"></a>[Excel для Windows](#tab/excel-windows)

Выполните следующую команду.

```
npm run start
```

Эта команда запускает веб-сервер и загружает неопубликованную надстройку пользовательской функции в Excel для Windows.

> [!NOTE]
> Если надстройка не загружается, проверьте правильность выполнения шага 3.

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

Выполните следующую команду.

```
npm run start-web
```

Эта команда запускает веб-сервер. Выполните шаги, описанные ниже, чтобы загрузить неопубликованную надстройку.

<ol type="a">
   <li>В Excel Online на вкладке <strong>Вставка</strong> выберите пункт <strong>Надстройки</strong>.<br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li>Выберите пункт <strong>Управление моими надстройками</strong>, а затем выберите <strong>Отправить мою надстройку</strong>.</li> 
   <li>Выберите <strong>Обзор... </strong> и откройте корневой каталог проекта, созданный генератором Yeoman.</li> 
   <li>Выберите файл <strong>manifest.xml</strong> и нажмите <strong>Открыть</strong>, а затем выберите <strong>Отправить</strong>.</li>
</ol>

> [!NOTE]
> Если надстройка не загружается, проверьте правильность выполнения шага 3.

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>Проверка работы готовой пользовательской функции

В проекте пользовательской функции, который вы создали, уже имеются две готовые пользовательские функции с именами ADD (Добавить) и INCREMENT (Увеличить). Код для этих встроенных функций содержится в файле **src/customfunctions.js**. Файл **./manifest.xml** указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`. Вы будете использовать пространство имен CONTOSO для доступа к пользовательским функциям в Excel.

Затем вы проверите пользовательскую функцию `ADD`, выполнив описанные ниже действия:

1. В Excel перейдите в любую ячейку и введите `=CONTOSO`. Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.

2. Выполните запуск функции `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.

Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете и возвращает результат **210**.

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Создание пользовательской функции, которая запрашивает данные из сети Интернет

Интеграция данных из Интернета — отличный способ расширения функционала Excel через пользовательские функции. Далее необходимо создать пользовательскую функцию под именем `stockPrice`, которая получает котировки акций из Web API и возвращает результат в ячейку на листе. Вы будете использовать API IEX Trading, который предоставляется бесплатно и не требует проверки подлинности.

1. В проекте **stock-ticker** найдите файл **src/customfunctions.js** и откройте его в редакторе кода.

2. В **customfunctions.js** найдите функцию `increment` и добавьте приведенный ниже код сразу после этой функции.

    ```js
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }

> [!NOTE]
> In the January Insiders 1901 Build, there is a bug preventing fetch calls from executing which will result in #VALUE!.
> To workaround this please use the [XMLHTTPRequest API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime#requesting-external-data) to make the web request.

3. In **customfunctions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    Код `CustomFunctions.associate` сопоставляет `id` функции с адресом функции `increment` в JavaScript, чтобы Excel мог вызвать вашу функцию.

    Прежде чем Excel сможет использовать вашу пользовательскую функцию, необходимо описать ее с помощью метаданных. Вам нужно определить `id`, используемый в методе `associate` ранее, а также некоторые другие метаданные.


4. Откройте файл **config/customfunctions.json**. Добавьте указанный ниже объект JSON в массив 'functions' и сохраните файл.

    ```JSON
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    Этот объект JSON описывает функцию `stockPrice`, ее параметры и тип результатов, который она возвращает.

5. Повторно зарегистрируйте надстройку в Excel, чтобы новая функция стала доступной. 

# <a name="excel-for-windowstabexcel-windows"></a>[Excel для Windows](#tab/excel-windows)

1. Закройте Excel, а затем откройте Excel повторно.

2. В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).

3. В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **stock-ticker**, чтобы зарегистрировать ее.
    ![Вставка ленты в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)

2. Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**. 

3. Выберите **Обзор... ** и откройте корневой каталог проекта, созданный генератором Yeoman. 

4. Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.

--- 

<ol start="6">
<li> Теперь давайте оценим, как работает новая функция. В ячейке <strong>B1</strong> введите нужный текст <strong>= CONTOSO. STOCKPRICE("MSFT")</strong> и нажмите ВВОД. Вы должны увидеть, что результат в ячейке <strong>B1</strong> является текущей ценой одной акции корпорации Майкрософт.</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>Создание потоковой асинхронной пользовательской функции

Функция `stockPrice` возвращает цену акции в конкретный момент времени, однако цены на акции всегда меняются. Далее вы создадите пользовательскую функцию с именем `stockPriceStream`, которая получает цену акции каждые 1000 милисекунд.

1. В проекте **stock-ticker** добавьте указанный ниже код в файл **src/customfunctions.js** и сохраните его.

    ```js
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    Прежде чем Excel сможет использовать вашу пользовательскую функцию, необходимо описать ее с помощью метаданных.
    
2. В проекте **stock-ticker** добавьте указанный ниже объект в массив `functions` в файле **config/customfunctions.json** и сохраните файл.
    
    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

    Объект JSON описывает функцию `stockPriceStream`. Для любой функции потоковой передачи свойство `stream` и свойство `cancelable` должны быть заданы как `true` в объекте `options`, как показано в этом примере кода.

3. Повторно зарегистрируйте надстройку в Excel, чтобы новая функция стала доступной.

# <a name="excel-for-windowstabexcel-windows"></a>[Excel для Windows](#tab/excel-windows)

1. Закройте Excel, а затем откройте Excel повторно.

2. В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).

3. В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **stock-ticker**, чтобы зарегистрировать ее.
    ![Вставка ленты в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)

2. Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.

3. Выберите **Обзор... ** и откройте корневой каталог проекта, созданный генератором Yeoman.

4. Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.

--- 

<ol start="4">
<li>Теперь давайте оценим, как работает новая функция. В ячейке <strong>C1</strong> введите нужный текст <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> и нажмите ВВОД. Если рынок ценных бумаг открыт, вы увидите, что результат в ячейке <strong>C1</strong> постоянно обновляется, отражая в режиме реального времени цену одной акции корпорации Майкрософт.</li>
</ol>

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем! Вы создали новый проект пользовательских функций, попробовали, как работает готовая функция, создали пользовательскую функцию, которая запрашивает данные из Интернета, а также создали пользовательскую функцию, которая осуществляет потоковую передачу данных в реальном времени из сети Интернет. Чтобы узнать больше о пользовательских функции в Excel, перейдите к следующей статье:

> [!div class="nextstepaction"]
> [Создание пользовательских функций в Excel](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>Юридические сведения

Данные предоставлены бесплатно компанией [IEX](https://iextrading.com/developer/). Ознакомьтесь с [Условиями использования IEX](https://iextrading.com/api-exhibit-a/). Корпорация Майкрософт использует API компании IEX в этом руководстве исключительно в ознакомительных целях.


