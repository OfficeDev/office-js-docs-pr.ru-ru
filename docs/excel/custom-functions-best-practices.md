---
ms.date: 09/20/2018
description: Рекомендации и рекомендуемые шаблоны для настраиваемых функций Excel.
title: Рекомендации по настраиваемым функциям
ms.openlocfilehash: 3934910c397aea348c4fe2d7f95f1dc20ebeb4d3
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985790"
---
# <a name="custom-functions-best-practices"></a>Рекомендации по настраиваемым функциям

В этой статье описаны рекомендации по разработке настраиваемых функций в Excel.

## <a name="error-handling"></a>Обработка ошибок

При построении надстройки, которая определяет настраиваемые функции, не забудьте включить логику обработки ошибок, возникающих в среде выполнения. Обработка ошибок для настраиваемых функций такая же, как и в случае [обработки ошибок для API JavaScript Excel в целом](excel-add-ins-error-handling.md). В следующем примере кода `.catch` будут обрабатываться все ошибки, возникающие ранее в коде.

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

- Используйте [ведение журнала времени выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) для отладки XML-файла манифеста вашей надстройки. 

- Используйте `console.log` операторы в коде настраиваемых функций для отправки выходных данных в консоль в режиме реального времени.

> [!NOTE]
> В настоящее время регистрация времени выполнения доступна только для классических приложений Office 2016.

## <a name="debugging"></a>Отладка

На данный момент наилучшим методом отладки пользовательских функций Excel является предварительная [загрузка неопубликованной надстройки](../testing/sideload-office-add-ins-for-testing.md) в Excel Online. Затем вы можете выполнить отладку настраиваемых функций с помощью [собственного средства отладки F12 вашего браузера](../testing/debug-add-ins-in-office-online.md).

Если надстройку не удалось зарегистрировать, [проверьте правильность настройки сертификатов SSL](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) для веб-сервера, где размещено приложение надстройки.

## <a name="mapping-names"></a>Сопоставление имен

По умолчанию, имя настраиваемой функции в файл JavaScript обычно объявляется полностью с помощьюпрописных букв и в точности соответствует имени функции, которую видят конечные пользователи в Excel. Тем не менее, можно изменить это с помощью `CustomFunctionsMappings` объекта для сопоставления одного или нескольких имен функции из файла JavaScript с разными значениями, которые  конечные пользователи увидят как имена функций в Excel. Эта функция полезна, если вы используете синтаксис методов uglifier, webpack или import - у каждого из которых есть трудности с именами функций в верхнем регистре. `CustomFunctionsMappings` возможно не обязателен для проектов, использующих JavaScript, но этот объект необходимо использовать, если в вашем проекте применяется TypeScript.  
  
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
