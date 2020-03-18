---
ms.date: 07/10/2019
description: Используйте `OfficeRuntime.storage`, чтобы сохранить состояние с пользовательскими функциями.
title: Сохранить и передать состояние в пользовательские функции.
localization_priority: Normal
ms.openlocfilehash: 8b55bfe61595b91f01c587282dc3f34887ce50fb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717203"
---
# <a name="save-and-share-state-in-custom-functions"></a>Сохранить и передать состояние в пользовательские функции.

Используйте объект `OfficeRuntime.storage`, чтобы сохранить состояние, связанное с пользовательскими функциями, или областью задач в вашей надстройке. Хранилище ограничено объемом 10 МБ на домен (который можно совместно использовать во множественных надстройках). В Excel для Windows объект `storage` представляет собой отдельное местоположение в пределах среды выполнения пользовательских функций, но в интернет-версии Excel и Excel для Mac объект `storage` тот же, что и `localStorage` браузера.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Есть несколько способов использовать `storage` для управления состоянием:

- Можно сохранить значения по умолчанию для пользовательских функций, чтобы применять их офлайн, когда нет доступа к веб-ресурсу.
- Можно сохранить значения для пользовательских функций, чтобы избежать дополнительных обращений к веб-ресурсу.
- Можно сохранить значения из пользовательской функции.
- Можно сохранить значения из области задач.

Показанный ниже пример кода иллюстрирует, как сохранить элемент в `storage` и получить его.

```js
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}
```

[Более подробный пример кода в GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) иллюстрирует, как передать эту информацию в область задач.

>[!NOTE]
> Объект `storage` заменяет собой предыдущий объект хранилища с именем `AsyncStorage`, которой сейчас не рекомендуется к использованию. Если используется объект `AsyncStorage` в текущем коде пользовательских функций, обновите его, чтобы использовать объект `storage`.

## <a name="next-steps"></a>Дальнейшие действия
Узнайте, как [автоматически генерировать метаданные JSON для своих пользовательских функций](custom-functions-json-autogeneration.md). 

## <a name="see-also"></a>См. также

* [Метаданные пользовательских функций](custom-functions-json.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Отладка пользовательских функций](custom-functions-debugging.md)
