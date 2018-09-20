# <a name="custom-function-metadata"></a>Метаданные настраиваемой функции

Когда вы включаете [настраиваемые функции](custom-functions-overview.md) в надстройке Excel, вы должны разместить файл JSON, содержащий метаданные о функциях (в дополнение к размещению файла JavaScript с функциями и HTML-файлом без пользовательского интерфейса, который будет служить родителем файла JavaScript). В этой статье описывается формат файла JSON с примерами.

Полная выборка файла JSON доступна [здесь](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions-array"></a>Массив функций

Метаданные — это объект JSON, содержащий одно свойство `functions`, значение которого представляет собой массив объектов. Каждый из этих объектов представляет собой одну настраиваемую функцию. Следующая таблица содержит ее свойства:

|  Свойство  |  Тип данных  |  Обязательность  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  строка  |  Нет  |  Описание функции, которая появляется в пользовательском интерфейсе Excel. Например, «Преобразует значение Цельсия в Фаренгейт». |
|  `helpUrl`  |  строка  |   Нет  |  URL-адрес, где ваши пользователи могут получить помощь по функции. (Он отображается в панели задач.) Например, «http://contoso.com/help/convertcelsiustofahrenheit.html»  |
|  `name`  |  строка  |  Да  |  Имя функции, которая будет отображаться (добавлено в пространстве имен) в пользовательском интерфейсе Excel, когда пользователь выбирает функцию. Оно должно совпадать с именем функции, указанном при ее определении в JavaScript. |
|  `options`  |  объект  |  Нет  |  Настройте, как Excel будет обрабатывать эту функцию. См. [объект опций](#options-object) для получения сведений. |
|  `parameters`  |  array  |  Да  |  Метаданные о параметрах функции. См. [массив параметров](#parameters-array) для получения сведений. |
|  `result`  |  объект  |  Да  |  Метаданные о значении, возвращаемом функцией. См. [объект результата](#result-object) для получения сведений. |

## <a name="options-object"></a>Объект Options

Объект `options` настраивает, как Excel обрабатывает эту функцию. Следующая таблица содержит ее свойства:

|  Свойство  |  Тип данных  |  Обязательность  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Нет, значение по умолчанию — `false`.  |  Если `true` Excel вызывает `onCanceled`обработчик каждый раз, когда пользователь предпринимает действие, которое имеет тот же эффект, что и отмена функции; например, вручную запуск пересчета или редактирования ячейки, на который ссылается функция. Если используется этот параметр, Excel вызовет функцию JavaScript с дополнительным `caller` параметром. (Не ***регистрируйте***этот параметр в `parameters`свойстве). В теле функции, должен быть назначен обработчик `caller.onCanceled`члена.|
|  `stream`  |  boolean  |  Нет, значение по умолчанию — `false`.  |  Если `true`, функция может выводить несколько раз в ячейку даже при вызове только один раз. Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`. (***Не***регистрируйте этот параметр в свойстве `parameters`.) Функция должна иметь выписку `return`. Вместо этого значение результата передается как аргумент метода `caller.setResult` обратного вызова.|

## <a name="parameters-array"></a>Массив параметров

Свойство `parameters` находится в массиве параметров. Каждый из этих объектов представляет собой параметр. Следующая таблица содержит ее свойства:

|  Свойство  |  Тип данных  |  Обязательность  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  строка  |  Нет |  Описание параметра.  |
|  `dimensionality`  |  строка  |  Да  |  Должно быть либо «скалярным», то есть значением без массива, либо «матрицей», то есть массивом массивов строк.  |
|  `name`  |  строка  |  Да  |  Имя параметра. Это имя отображается в Excel IntelliSense.  |
|  `type`  |  строка  |  Да  |  Тип данных параметра. Должно быть «логический», «числовой» или «строка».  |

## <a name="result-object"></a>Результирующий объект

Свойство `results` предоставляет метаданные о значении, возвращаемом функцией. Следующая таблица содержит ее свойства:

|  Свойство  |  Тип данных  |  Обязательность  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  строка  |  Нет  |  Должно быть либо «скалярным», то есть значением без массива, либо «матрицей», то есть массивом массивов строк.  |
|  `type`  |  строка  |  Да  |  Тип данных параметра. Должно быть «логический», «числовой» или «строка».  |

## <a name="example"></a>Пример

Следующий код JSON является примером файла метаданных для пользовательских функций.

```json
{
    "functions": [
        {
            "name": "ADD42", 
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "name": "ADD42ASYNC", 
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "increment",
                    "description": "the number to be added each time",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "stream": true,
                "cancelable": true
            }
        },
        {
            "name": "SECONDHIGHEST", 
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "range",
                    "description": "the input range",
                    "type": "number",
                    "dimensionality": "matrix"
                }
            ]
        }
    ]
}

```

## <a name="see-also"></a>См. также
[Настраиваемые функции](custom-functions-overview.md)<br>
[Руководства и примеры формул массива](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
