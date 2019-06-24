---
ms.date: 06/18/2019
description: Обработка ошибок в пользовательских функциях Excel.
title: Обработка ошибок в пользовательских функциях Excel
localization_priority: Priority
ms.openlocfilehash: 3818d33121ed26bb7d65c56bf6c504f2fb049c72
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127921"
---
# <a name="error-handling-within-custom-functions"></a>Обработка ошибок в пользовательских функциях

При создании надстройки, которая определяет пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Обработка ошибок для пользовательских функций в значительной степени совпадает с [обработкой ошибок для API JavaScript в Excel](excel-add-ins-error-handling.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

В следующем примере кода `.catch` будет обрабатывать любые ошибки, возникающие ранее в коде.

```js
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="next-steps"></a>Дальнейшие действия
Узнайте, как [устранять проблемы с пользовательскими функциями](custom-functions-troubleshooting.md).

## <a name="see-also"></a>Дополнительные ресурсы

* [Отладка пользовательских функций](custom-functions-debugging.md)
* [Требования к настраиваемым функциям](custom-functions-requirements.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
