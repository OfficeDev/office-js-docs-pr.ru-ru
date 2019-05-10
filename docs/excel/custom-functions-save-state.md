---
ms.date: 05/03/2019
description: Используйте `OfficeRuntime.storage`, чтобы сохранить состояние с пользовательскими функциями.
title: Сохранить и передать состояние в пользовательские функции.
localization_priority: Priority
ms.openlocfilehash: e7229637705425198d3905c1ff4102022fb99f1e
ms.sourcegitcommit: 95cf577178c1c88ebbffa77367634658eb7fe925
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2019
ms.locfileid: "33643832"
---
## <a name="save-and-share-state-in-custom-functions"></a>Сохранить и передать состояние в пользовательские функции.

Используйте объект `OfficeRuntime.storage`, чтобы сохранить состояние, связанное с пользовательскими функциями, или областью задач в вашей надстройке. Хранилище ограничено объемом 10 МБ на домен (который можно совместно использовать во множественных надстройках). В Excel для Windows объект `storage` представляет собой отдельное местоположение в пределах среды выполнения пользовательских функций, но в Excel Online и Excel для Mac объект `storage` тот же, что и `localStorage` браузера.

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

CustomFunctions.associate("STOREVALUE", StoreValue);
CustomFunctions.associate("GETVALUE", GetValue);
```

[Более подробный пример кода в GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) иллюстрирует, как передать эту информацию в область задач.

>[!NOTE]
> Объект `storage` заменяет собой предыдущий объект хранилища с именем `AsyncStorage`, которой сейчас не рекомендуется к использованию. Если используется объект `AsyncStorage` в текущем коде пользовательских функций, обновите его, чтобы использовать объект `storage`.

## <a name="next-steps"></a>Дальнейшие действия
Узнайте, как [автоматически генерировать метаданные JSON для своих пользовательских функций](custom-functions-json-autogeneration.md). 

## <a name="see-also"></a>См. также

* [Метаданные пользовательских функций](custom-functions-json.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Отладка пользовательских функций](custom-functions-debugging.md)
