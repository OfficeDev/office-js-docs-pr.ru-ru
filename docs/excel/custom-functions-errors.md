---
ms.date: 02/08/2019
description: Обработка ошибок в пользовательских функциях Excel.
title: Обработка ошибок в пользовательских функциях Excel (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: 170da03331663d6779bed7bf0bf5a9b75b908b3f
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632697"
---
# <a name="error-handling-within-custom-functions"></a>Обработка ошибок в пользовательских функциях

При создании надстройки, которая определяет пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Обработка ошибок для пользовательских функций в значительной степени совпадает с [обработкой ошибок для API JavaScript в Excel](excel-add-ins-error-handling.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

В следующем примере кода `.catch` будет обрабатывать любые ошибки, возникающие ранее в коде.

```js
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

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

## <a name="see-also"></a>См. также

* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Рекомендации по пользовательским функциям](custom-functions-best-practices.md)
* [Журнал изменений пользовательских функций](custom-functions-changelog.md)
