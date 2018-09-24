---
ms.date: 09/20/2018
description: Рекомендации и рекомендуемые шаблоны для настраиваемых функций Excel.
title: Рекомендации по настраиваемым функциям
ms.openlocfilehash: 1f2c0a80e62b65523fcc1673ba2ca4be444e6ce0
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068827"
---
# <a name="custom-functions-best-practices"></a>Рекомендации по настраиваемым функциям

В этой статье описаны рекомендации по разработке настраиваемых функций в Excel.

## <a name="error-handling"></a>Обработка ошибок

При построении надстройки, которая определяет настраиваемые функции, не забудьте включить логику обработки ошибок для учетной записи для среды выполнения ошибок. Обработка ошибок для настраиваемых функций совпадает с [обработкой ошибок для Excel API JavaScript в целом](excel-add-ins-error-handling.md). В следующем примере кода `.catch` будут обрабатываться все ошибки, возникающие ранее в коде.

```js
function getComment(x) {
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

## <a name="error-logging"></a>Ведение журнала ошибок

Можно включить журнал ведения  ошибки для настраиваемых функций надстройки несколькими способами, такими как: 

- [Используйте регистрациию времени выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) для отладки надстройки в XML-файле манифеста. 

- Используйте `console.log` операторы в коде настраиваемых функций для отправки выходных данных в консоль в режиме реального времени.

> [!NOTE]
> В настоящее время  регистрация времени выполнения доступна только для рабочего стола Office 2016.

## <a name="debugging"></a>Отладка

На данный момент наилучшим методом для отладки настраиваемых функций Excel является использование [Excel Online](https://www.office.com/launch/excel) и использование средства отладки F12, встроенного в ваш браузер. Дополнительные средства отладки для настраиваемых функций могут быть доступны в будущем.

## <a name="mapping-names"></a>Сопоставление имен

По умолчанию, имя настраиваемой функции в файл JavaScript обычно объявляется полностью с помощьюпрописных букв и в точности соответствует имени функции, которую видят конечные пользователи в Excel. Тем не менее, можно изменить это с помощью `CustomFunctionsMappings` объекта для сопоставления одного или нескольких имен функции из файла JavaScript с разными значениями, которые  конечные пользователи увидят как имена функций в Excel. Несмотря на то, что необходимо использовать `CustomFunctionsMapping`, это может быть полезно, если вы используете  синтаксис uglifier, webpack или import - каждый из которых имеет трудности с прописными буквами имен функции.
  
В следующем примере кода определяется одна пара "ключ-значение", которая сопоставляет имя функции JavaScript `plusFortyTwo` с `ADD42` именем функции в пользовательском Интерфейсе Excel. Когда конечный пользователь выбирает `ADD42` функцию в Excel, `plusFortyTwo`запускается функция JavaScript.

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

В следующем примере кода определяются две пары "ключ-значение". Первая пара сопоставляет имя функции JavaScript `plusFifty` с `ADD50` именем функции в пользовательском Интерфейсе Excel и вторая пара сопоставляет имя функции JavaScript `plusOneHundred` с `ADD100` именем функции в пользовательском Интерфейсе Excel. Когда конечный пользователь выбирает `ADD50` функцию в Excel, `plusFifty`запускается функция JavaScript. Когда конечный пользователь выбирает `ADD100` функцию в Excel, `plusOneHundred`запускается функция JavaScript.

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## <a name="see-also"></a>См. также

* [Создание настраиваемых функций в Excel](custom-functions-overview.md)
* [Настраиваемые функции метаданных](custom-functions-json.md)
* [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md)