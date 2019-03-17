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
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="7debb-103">Обработка ошибок в пользовательских функциях</span><span class="sxs-lookup"><span data-stu-id="7debb-103">Error handling within custom functions</span></span>

<span data-ttu-id="7debb-104">При создании надстройки, которая определяет пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="7debb-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="7debb-105">Обработка ошибок для пользовательских функций в значительной степени совпадает с [обработкой ошибок для API JavaScript в Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="7debb-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="7debb-106">В следующем примере кода `.catch` будет обрабатывать любые ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="7debb-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="7debb-107">См. также</span><span class="sxs-lookup"><span data-stu-id="7debb-107">See also</span></span>

* [<span data-ttu-id="7debb-108">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="7debb-108">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="7debb-109">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="7debb-109">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="7debb-110">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="7debb-110">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="7debb-111">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="7debb-111">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="7debb-112">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="7debb-112">Custom functions changelog</span></span>](custom-functions-changelog.md)
