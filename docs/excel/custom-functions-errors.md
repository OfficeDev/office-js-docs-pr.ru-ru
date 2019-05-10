---
ms.date: 05/03/2019
description: Обработка ошибок в пользовательских функциях Excel.
title: Обработка ошибок в пользовательских функциях Excel
localization_priority: Priority
ms.openlocfilehash: 188ece6c77bc2cafad6f22448fb698e0c0370ef8
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628160"
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
