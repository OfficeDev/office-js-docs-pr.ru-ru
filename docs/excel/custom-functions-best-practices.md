---
ms.date: 05/06/2019
description: Ознакомьтесь с рекомендациями по разработке пользовательских функций в Excel.
title: Рекомендации в отношении пользовательских функций
localization_priority: Normal
ms.openlocfilehash: 7369faa463966dd309258bf431eae8719407be38
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628149"
---
# <a name="custom-functions-best-practices"></a>Рекомендации в отношении пользовательских функций

В этой статье описаны рекомендации по разработке пользовательских функций в Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a>Сопоставление имен функций с метаданными JSON

Как описано в статье [Обзор пользовательских функций](custom-functions-overview.md) проект пользовательских функций должен содержать как файл метаданных JSON, так и файл сценария (JavaScript или TypeScript) для образования готовой функции. Если вы используете `yo office` метаданные JSON, можно создавать их из комментариев кода. В противном случае вам потребуется создать файл метаданных JSON вручную.

Чтобы функция работала должным образом, необходимо связать `id` свойство функции с реализацией JavaScript. Убедитесь, что существует связь, иначе функция не будет вызываться. В приведенном ниже примере кода показано, как выполнить связь с `CustomFunctions.associate()` помощью метода. Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

В следующем JSON показаны метаданные JSON, связанные с предыдущим кодом пользовательской функции JavaScript.

```json
{
  "functions": [
    {
        "description": "Add two numbers",
        "id": "ADD",
        "name": "ADD",
        "parameters": [
            {
                "description": "First number",
                "name": "first",
                "type": "number"
            },
            {
                "description": "Second number",
                "name": "second",
                "type": "number"
            }
        ],
        "result": {
            "type": "number"
        }
    },
  ]
}
```


Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.

* Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.

* Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла. То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.

* Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript. Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.

* В файле JavaScript укажите настраиваемое сопоставление функций с помощью `CustomFunctions.associate` каждой функции.

В приведенном ниже примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript. Значения `id` свойств `name` и представлены в верхнем регистре, что является лучшим вариантом при описании пользовательских функций. Этот код JSON необходимо добавить только в том случае, если вы готовите собственный файл JSON вручную и не используете автоматическое создание. Для получения дополнительной информации об автоформировании, ознакомьтесь со статьей [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="additional-considerations"></a>Дополнительные рекомендации

Избегайте прямого или косвенного доступа к объектной модели документов (DOM) (например, с помощью jQuery) из пользовательской функции. В Excel для Windows, где пользовательские функции используют [среду выполнения JavaScript](custom-functions-runtime.md), пользовательские функции не могут выполнять доступ к DOM.

## <a name="next-steps"></a>Дальнейшие действия
Узнайте, как [выполнять веб-запросы с пользовательскими функциями](custom-functions-web-reqs.md).

## <a name="see-also"></a>См. также

* [Автоматически создавать метаданные JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
