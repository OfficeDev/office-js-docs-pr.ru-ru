---
ms.date: 05/17/2020
description: Общие сведения о пользовательских функциях Excel, не использующих область задач и определенную среду выполнения JavaScript.
title: Среда выполнения для пользовательских функций Excel без пользовательского интерфейса
localization_priority: Normal
ms.openlocfilehash: 5cb9aa480d6923d31434d58a9683e9a9f5d48458
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609645"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>Среда выполнения для пользовательских функций Excel без пользовательского интерфейса

Пользовательские функции, которые не используют область задач (пользовательские функции без пользовательского интерфейса), используют среду выполнения JavaScript, предназначенную для оптимизации производительности вычислений.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Эта среда выполнения JavaScript предоставляет доступ к API в `OfficeRuntime` пространстве имен, которые можно использовать в пользовательских функциях без пользовательского интерфейса, и область задач для хранения данных.

## <a name="requesting-external-data"></a>Запрос внешних данных

В пользовательской функции без пользовательского интерфейса можно запрашивать внешние данные с помощью API, например [Извлечение](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) или с помощью [XMLHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.

Имейте в виду, что функции без пользовательского интерфейса должны использовать дополнительные меры безопасности при создании XmlHttpRequest, для чего требуется [одна и та же политика начала](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простая [CORS](https://www.w3.org/TR/cors/).

Простая реализация CORS не может использовать файлы cookie и поддерживает только простые методы (GET, HEAD, POST). Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`. Вы также можете использовать `Content-Type` заголовок в простой CORS, при условии, что тип контента: `application/x-www-form-urlencoded` , `text/plain` или `multipart/form-data` .

## <a name="storing-and-accessing-data"></a>Хранения данных и доступ к ним

В пользовательской функции без пользовательского интерфейса можно хранить и получать доступ к данным с помощью `OfficeRuntime.storage` объекта. `Storage`— Это постоянная, незашифрованная система хранения с ключом, которая предоставляет альтернативу [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), который не может использоваться пользовательскими функциями без пользовательского интерфейса. `Storage`предоставляет 10 МБ данных для каждого домена. Домены могут совместно использоваться несколькими надстройками.

`Storage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным. Например, маркеры для проверки подлинности пользователей могут храниться в `storage` связи с тем, что они доступны как в пользовательских функциях, так и в пользовательских элементах пользовательского интерфейса, таких как область задач. Аналогично, если две надстройки совместно используют один и тот же домен (например, `www.contoso.com/addin1` `www.contoso.com/addin2` ), они также могут обмениваться данными с помощью `storage` . Обратите внимание, что надстройки, у которых есть разные поддомены, будут иметь разные экземпляры `storage` (например, `subdomain.contoso.com/addin1` `differentsubdomain.contoso.com/addin2` ).

Так как `storage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".

Ниже указаны методы, доступные в объекте `storage`.

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

.[!NOTE]
> Нет метода для очистки всей информации (например, `clear` ). Вместо этого вам следует использовать `removeItems` для одновременного удаления нескольких записей.

### <a name="officeruntimestorage-example"></a>Пример Оффицерунтиме. Storage

В следующем примере кода вызывается `OfficeRuntime.storage.setItem` функция для установки ключа и значения `storage` .

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>Дополнительные рекомендации

Если надстройка использует только пользовательские функции без пользовательского интерфейса, обратите внимание на то, что вы не можете получить доступ к объектной модели документов (DOM) с пользовательскими функциями без пользовательского интерфейса или использовать библиотеки, такие как jQuery, которые используют модель DOM.

## <a name="next-steps"></a>Дальнейшие действия
Узнайте, как [отлаживать пользовательские функции без пользовательского интерфейса](custom-functions-debugging.md).

## <a name="see-also"></a>См. также

* [Проверка подлинности пользовательских функций без пользовательского интерфейса](custom-functions-authentication.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md)
