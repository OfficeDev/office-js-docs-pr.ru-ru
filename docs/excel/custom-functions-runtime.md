---
ms.date: 06/15/2022
description: Сведения Excel пользовательских функций, которые не используют общую среду выполнения и определенную среду выполнения JavaScript.
title: Среда выполнения только для JavaScript для пользовательских функций
ms.localizationpriority: medium
ms.openlocfilehash: 614e96937c769307b58e66943caa499f1f12d92c
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229668"
---
# <a name="javascript-only-runtime-for-custom-functions"></a>Среда выполнения только для JavaScript для пользовательских функций

Пользовательские функции, не использующие общую среду выполнения, используют среду выполнения только для JavaScript, предназначенную для оптимизации производительности вычислений.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Эта среда выполнения JavaScript предоставляет доступ к интерфейсам API `OfficeRuntime` в пространстве имен, которые могут использоваться пользовательскими функциями, и области задач (которая выполняется в другой среде выполнения) для хранения данных.

## <a name="request-external-data"></a>Запрос внешних данных

В пределах пользовательской функции можно запрашивать внешние данные с помощью такого API, как [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API), или с помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.

Имейте в виду, что пользовательские функции должны использовать дополнительные меры безопасности при создании XmlHttpRequests, для которых требуется та же политика источника и [простая CORS](https://www.w3.org/TR/cors/).[](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)

Простая реализация CORS не может использовать файлы cookie и поддерживает только простые методы (GET, HEAD, POST). Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`. Вы также можете использовать заголовок `Content-Type` в простом CORS при условии, что тип контента — `application/x-www-form-urlencoded`или `multipart/form-data``text/plain`.

## <a name="store-and-access-data"></a>Хранение данных и доступ к ним

В пользовательской функции, которая не использует общую среду выполнения, можно хранить данные и получать к ним доступ с помощью [объекта OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) . Объект `Storage` представляет собой постоянную незашифрованную систему хранения "ключ-значение", которая предоставляет альтернативу [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), которая не может использоваться пользовательскими функциями, использующими среду выполнения только для JavaScript. Объект `Storage` предоставляет 10 МБ данных на домен. Домены могут совместно использоваться несколькими надстройки.

Объект `Storage` является общим решением для хранения, то есть несколько частей надстройки могут получить доступ к одинаковым данным. Например, `Storage` маркеры для проверки подлинности пользователей могут храниться в объекте, так как к ним можно получить доступ как с помощью пользовательской функции (с помощью среды выполнения только для JavaScript), так и с помощью области задач (с использованием полной среды выполнения веб-представления). Аналогичным образом, если две надстройки совместно используют один и тот же домен (например, `www.contoso.com/addin1`), `www.contoso.com/addin2``Storage` они также могут предоставлять общий доступ к информации через объект. Обратите внимание, что надстройки `Storage` с разными поддоменами будут иметь разные экземпляры (например, `subdomain.contoso.com/addin1``differentsubdomain.contoso.com/addin2`

Поскольку объект `Storage` может быть общим расположением, важно понимать, что можно переопределить пары "ключ-значение".

Для объекта доступны следующие `Storage` методы.

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> Нет метода очистки всех сведений (например, `clear`). Вместо этого вам следует использовать `removeItems` для одновременного удаления нескольких записей.

### <a name="officeruntimestorage-example"></a>Пример OfficeRuntime.storage

В следующем примере кода функция `OfficeRuntime.storage.setItem` вызывается для задания ключа и значения `storage`в .

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="next-steps"></a>Дальнейшие действия

Узнайте, как [выполнять отладку пользовательских функций](custom-functions-debugging.md).

## <a name="see-also"></a>См. также

* [Проверка подлинности для пользовательских функций без общей среды выполнения](custom-functions-authentication.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md)
