---
ms.date: 07/08/2021
description: Понимание Excel пользовательских функций, которые не используют области задач и их определенное время запуска JavaScript.
title: Время запуска для пользовательских Excel пользовательских функций
ms.localizationpriority: medium
ms.openlocfilehash: 491e47674d87d99d0adeda952ee65ffc24dff2bd
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150693"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>Время запуска для пользовательских Excel пользовательских функций

Настраиваемые функции, которые не используют области задач (не пользовательские функции без пользовательского интерфейса), используют время выполнения JavaScript, предназначенное для оптимизации производительности вычислений.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Это время запуска JavaScript предоставляет доступ к API в пространстве имен, которые могут использоваться пользовательскими функциями без пользовательского интерфейса и областью задач для `OfficeRuntime` хранения данных.

## <a name="request-external-data"></a>Запрос внешних данных

В настраиваемой функции без пользовательского интерфейса можно запрашивать внешние данные с помощью API типа [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) или С помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)— стандартного веб-API, который выдает HTTP-запросы для взаимодействия с серверами.

Следует помнить, что функции, не требующие пользовательского интерфейса, должны [](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) применять дополнительные меры безопасности при создании XmlHttpRequests, для которых требуется та же политика происхождения и [простая CORS.](https://www.w3.org/TR/cors/)

Простая реализация CORS не может использовать файлы cookie и поддерживает только простые методы (GET, HEAD, POST). Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`. Вы также можете использовать `Content-Type` заготку в простой CORS, при условии, что тип контента `application/x-www-form-urlencoded` , `text/plain` или `multipart/form-data` .

## <a name="store-and-access-data"></a>Хранение и доступ к данным

В настраиваемой функции без пользовательского интерфейса можно хранить и получать доступ к данным с помощью `OfficeRuntime.storage` объекта. `Storage` — это система хранения сохраняемой, незашифрованной и ключевой ценности, которая предоставляет альтернативу [localStorage,](https://developer.mozilla.org/docs/Web/API/Window/localStorage)которая не может использоваться пользовательскими функциями без пользовательского интерфейса. `Storage` предоставляет 10 МБ данных на домен. Домены могут быть общими для более чем одной надстройки.

`Storage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным. Например, маркеры для проверки подлинности пользователей могут храниться, так как к нему можно получить доступ как с помощью настраиваемой функции без пользовательского интерфейса, так и элементов пользовательского интерфейса надстройки, таких как области `storage` задач. Аналогичным образом, если два надстройки имеют один и тот же домен (например, ), им также разрешено обмениваться информацией взад и `www.contoso.com/addin1` `www.contoso.com/addin2` вперед через `storage` . Обратите внимание, что надстройки с разными поддоменами будут иметь различные экземпляры `storage` `subdomain.contoso.com/addin1` (например, , `differentsubdomain.contoso.com/addin2` ).

Так как `storage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".

На объекте доступны следующие `storage` методы.

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> Нет способа очистки всех сведений `clear` (например). Вместо этого вам следует использовать `removeItems` для одновременного удаления нескольких записей.

### <a name="officeruntimestorage-example"></a>Пример OfficeRuntime.storage

В следующем примере кода функция вызывает функцию `OfficeRuntime.storage.setItem` для набора ключа и значения `storage` в .

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

Если надстройка использует только настраиваемые функции без пользовательского интерфейса, обратите внимание, что вы не можете получить доступ к объектной модели документа (DOM) с пользовательскими функциями без пользовательского интерфейса или использовать библиотеки, такие как jQuery, которые полагаются на DOM.

## <a name="next-steps"></a>Следующие шаги

Узнайте, как [отламыть пользовательские функции без пользовательского интерфейса.](custom-functions-debugging.md)

## <a name="see-also"></a>Дополнительные материалы

* [Проверка подлинности пользовательских функций без пользовательского интерфейса](custom-functions-authentication.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md)
