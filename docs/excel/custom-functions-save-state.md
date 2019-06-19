---
ms.date: 06/17/2019
description: Используйте `OfficeRuntime.storage`, чтобы сохранить состояние с пользовательскими функциями.
title: Сохранить и передать состояние в пользовательские функции.
localization_priority: Priority
ms.openlocfilehash: c0825b67bfb97cea75e09704969e915d9560e39e
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059890"
---
# <a name="save-and-share-state-in-custom-functions"></a><span data-ttu-id="dfca8-103">Сохранить и передать состояние в пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="dfca8-103">Save and share state in custom functions</span></span>

<span data-ttu-id="dfca8-104">Используйте объект `OfficeRuntime.storage`, чтобы сохранить состояние, связанное с пользовательскими функциями, или областью задач в вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="dfca8-104">Use the `OfficeRuntime.storage` object to save state related to custom functions or the task pane in your add-in.</span></span> <span data-ttu-id="dfca8-105">Хранилище ограничено объемом 10 МБ на домен (который можно совместно использовать во множественных надстройках).</span><span class="sxs-lookup"><span data-stu-id="dfca8-105">Storage is limited to 10 MB per domain (which may be shared across multiple add-ins).</span></span> <span data-ttu-id="dfca8-106">В Excel для Windows объект `storage` представляет собой отдельное местоположение в пределах среды выполнения пользовательских функций, но в Excel Online и Excel для Mac объект `storage` тот же, что и `localStorage` браузера.</span><span class="sxs-lookup"><span data-stu-id="dfca8-106">In Excel on Windows, the `storage` object is a separate location within the custom functions runtime, but for Excel Online and Excel for Mac, the `storage` object is the same as the browser's `localStorage`.</span></span>

<span data-ttu-id="dfca8-107">Есть несколько способов использовать `storage` для управления состоянием:</span><span class="sxs-lookup"><span data-stu-id="dfca8-107">There are multiple ways to use `storage` for state management:</span></span>

- <span data-ttu-id="dfca8-108">Можно сохранить значения по умолчанию для пользовательских функций, чтобы применять их офлайн, когда нет доступа к веб-ресурсу.</span><span class="sxs-lookup"><span data-stu-id="dfca8-108">You can store default values for custom functions to use when you are offline and unable to reach a web resource.</span></span>
- <span data-ttu-id="dfca8-109">Можно сохранить значения для пользовательских функций, чтобы избежать дополнительных обращений к веб-ресурсу.</span><span class="sxs-lookup"><span data-stu-id="dfca8-109">You can save values for custom functions to use to avoid making additional calls to a web resource.</span></span>
- <span data-ttu-id="dfca8-110">Можно сохранить значения из пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="dfca8-110">You can save values from your custom function.</span></span>
- <span data-ttu-id="dfca8-111">Можно сохранить значения из области задач.</span><span class="sxs-lookup"><span data-stu-id="dfca8-111">You can store values from your task pane.</span></span>

<span data-ttu-id="dfca8-112">Показанный ниже пример кода иллюстрирует, как сохранить элемент в `storage` и получить его.</span><span class="sxs-lookup"><span data-stu-id="dfca8-112">The following code sample illustrates how to store an item into `storage` and retrieve it.</span></span>

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

<span data-ttu-id="dfca8-113">[Более подробный пример кода в GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) иллюстрирует, как передать эту информацию в область задач.</span><span class="sxs-lookup"><span data-stu-id="dfca8-113">[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane.</span></span>

>[!NOTE]
> <span data-ttu-id="dfca8-114">Объект `storage` заменяет собой предыдущий объект хранилища с именем `AsyncStorage`, которой сейчас не рекомендуется к использованию.</span><span class="sxs-lookup"><span data-stu-id="dfca8-114">The `storage` object replaces the previous storage object named `AsyncStorage` which is now deprecated.</span></span> <span data-ttu-id="dfca8-115">Если используется объект `AsyncStorage` в текущем коде пользовательских функций, обновите его, чтобы использовать объект `storage`.</span><span class="sxs-lookup"><span data-stu-id="dfca8-115">If using the `AsyncStorage` object in your current custom functions code, please update it to use the `storage` object.</span></span>

## <a name="next-steps"></a><span data-ttu-id="dfca8-116">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="dfca8-116">Next steps</span></span>
<span data-ttu-id="dfca8-117">Узнайте, как [автоматически генерировать метаданные JSON для своих пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="dfca8-117">Learn how to [autogenerate the JSON metadata for your custom functions](custom-functions-json-autogeneration.md).</span></span> 

## <a name="see-also"></a><span data-ttu-id="dfca8-118">См. также</span><span class="sxs-lookup"><span data-stu-id="dfca8-118">See also</span></span>

* [<span data-ttu-id="dfca8-119">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="dfca8-119">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="dfca8-120">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="dfca8-120">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="dfca8-121">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="dfca8-121">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="dfca8-122">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="dfca8-122">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="dfca8-123">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="dfca8-123">Custom functions debugging</span></span>](custom-functions-debugging.md)
